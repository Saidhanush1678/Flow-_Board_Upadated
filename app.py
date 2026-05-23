import os
import secrets
import random
import numpy as np
import pandas as pd
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, session, jsonify, abort
from flask_socketio import SocketIO, emit, join_room, leave_room
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

# =====================================================
# SYSTEM CONFIGURATION & ENGINES
# =====================================================
app = Flask(__name__)
# Render instances can restart; using a persistent environment fallback keeps sessions alive between restarts
app.secret_key = os.environ.get('SECRET_KEY', secrets.token_hex(24))
app.config['SECRET_KEY'] = app.secret_key

# Changed async_mode to "gevent" for robust production WebSocket worker pooling
socketio = SocketIO(app, cors_allowed_origins="*", async_mode="gevent")

EXCEL_DB = "flowboard_database.xlsx"
PROJECT_FOLDER = "team_codes"

if not os.path.exists(PROJECT_FOLDER):
    os.makedirs(PROJECT_FOLDER)

attendance = []
notifications = []
MEETING_LINK = "https://meet.google.com/wge-aofy-frc"
active_user_dashboards = {} 

# =====================================================
# UNIFIED EXCEL STORAGE ENGINE
# =====================================================
def init_excel_database():
    if os.path.exists(EXCEL_DB):
        try:
            with pd.ExcelWriter(EXCEL_DB, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                if 'Meetings' not in writer.book.sheetnames:
                    df_m = pd.DataFrame(columns=["meeting_id", "title", "host_email", "invitee_email", "scheduled_time", "room_token"])
                    df_m.to_excel(writer, sheet_name='Meetings', index=False)
        except Exception:
            pass
        return

    with pd.ExcelWriter(EXCEL_DB, engine='openpyxl') as writer:
        df_users = pd.DataFrame([
            {"name": "System Manager Admin", "email": "manager@gmail.com", "password": "manager123", "role": "Manager"},
            {"name": "Vaishu Malla", "email": "vaishusamalla11@gmail.com", "password": "1234", "role": "Team Leader"},
            {"name": "Sravani", "email": "lakkavatrisravani@gmail.com", "password": "1234", "role": "Team Leader"},
            {"name": "Abhi", "email": "emulaabhi@gmail.com", "password": "1234", "role": "Team Member"},
            {"name": "Lalasa", "email": "nandagirilalasa01@gmail.com", "password": "1234", "role": "Team Member"},
            {"name": "Client User", "email": "client@gmail.com", "password": "1234", "role": "Client"}
        ])
        df_users.to_excel(writer, sheet_name='SystemUsers', index=False)
        
        df_projects = pd.DataFrame([
            {"id": "P-101", "name": "Cloud Infrastructure Core", "description": "Global enterprise migration setup.", "deadline": "2026-12-01", "status": "In Progress", "assigned_leader_email": "vaishusamalla11@gmail.com", "progress": 65},
            {"id": "P-102", "name": "AI Pipeline Engine", "description": "Automated data learning pipelines.", "deadline": "2026-11-15", "status": "Pending", "assigned_leader_email": "", "progress": 0},
            {"id": "P-103", "name": "React Dashboard Interface", "description": "Frontend matrix workspace display.", "deadline": "2026-07-19", "status": "Completed", "assigned_leader_email": "lakkavatrisravani@gmail.com", "progress": 100},
            {"id": "P-104", "name": "Cybersecurity Gateway", "description": "Build a secure web platform app with payment gateway APIs.", "deadline": "2026-08-30", "status": "Pending", "assigned_leader_email": "", "progress": 12}
        ])
        df_projects.to_excel(writer, sheet_name='Projects', index=False)
        
        df_tasks = pd.DataFrame([
            {"project": "P-101", "module": "Backend API Node", "assigned_to": "emulaabhi@gmail.com", "status": "Assigned", "progress": 20, "team_leader": "vaishusamalla11@gmail.com", "filename": "P101_backend_workspace.txt"},
            {"project": "P-101", "module": "Frontend User UI", "assigned_to": "nandagirilalasa01@gmail.com", "status": "Assigned", "progress": 45, "team_leader": "vaishusamalla11@gmail.com", "filename": "P101_frontend_workspace.txt"}
        ])
        df_tasks.to_excel(writer, sheet_name='SubTasks', index=False)

        df_meetings = pd.DataFrame([
            {"meeting_id": "M-501", "title": "Architecture Alignment Sync", "host_email": "manager@gmail.com", "invitee_email": "client@gmail.com", "scheduled_time": "2026-06-10 14:00", "room_token": "room_manager_client"}
        ])
        df_meetings.to_excel(writer, sheet_name='Meetings', index=False)

def read_sheet(sheet_name):
    try:
        return pd.read_excel(EXCEL_DB, sheet_name=sheet_name).fillna("").to_dict(orient='records')
    except Exception:
        return []

def write_sheet(sheet_name, data_list):
    with pd.ExcelWriter(EXCEL_DB, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        pd.DataFrame(data_list).to_excel(writer, sheet_name=sheet_name, index=False)

init_excel_database()

# =====================================================
# UNIVERSAL SETTINGS ROUTE
# =====================================================
@app.route('/settings', methods=['GET', 'POST'])
def settings():
    if not session.get('user'):
        return redirect('/login')
    
    current_email = session.get('user')
    users = read_sheet('SystemUsers')
    user_record = next((u for u in users if str(u['email']).lower() == current_email.lower()), None)
    
    if request.method == 'POST':
        new_name = request.form.get('name', '').strip()
        session['theme'] = request.form.get('theme', 'dark')
        
        if user_record and new_name:
            for u in users:
                if str(u['email']).lower() == current_email.lower():
                    u['name'] = new_name
                    break
            write_sheet('SystemUsers', users)
            session['name'] = new_name
            notifications.append(f"Profile config modified for node: {current_email}.")
            
        if session.get('role') == 'Manager':
            return redirect('/manager_dashboard')
        elif session.get('role') == 'Team Leader':
            return redirect(url_for('teamleader', email=current_email))
        elif session.get('role') == 'Team Member':
            return redirect(url_for('teammember', email=current_email))
        elif session.get('role') == 'Client':
            return redirect('/client')
            
    return render_template("settings.html", user_record=user_record)

# =====================================================
# AUTHENTICATION ROUTING GATEWAY
# =====================================================
@app.route('/')
def home():
    return render_template("index.html")

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'GET':
        return render_template("login.html")

    form_email = request.form['email'].strip().lower()
    form_password = request.form['password'].strip()

    users = read_sheet('SystemUsers')
    for u in users:
        if str(u['email']).lower() == form_email and str(u['password']) == form_password:
            session["user"] = u["email"]
            session["role"] = u["role"]
            session["name"] = u["name"]

            # Track attendance entry manually upon successful profile validation
            attendance.append({"name": u["name"], "time": datetime.now().strftime("%H:%M:%S")})

            if u["role"] == "Manager":
                return redirect('/manager')
            elif u["role"] == "Team Leader":
                return redirect(url_for('teamleader', email=u['email']))
            elif u["role"] == "Team Member":
                return redirect(url_for('teammember', email=u['email']))
            elif u["role"] == "Client":
                return redirect('/client')

    return "Invalid System Authentication Credentials Provided."

# =====================================================
# OPERATIONAL MANAGEMENT PANELS & CODE EDIT PRIVILEGES
# =====================================================
@app.route('/manager')
@app.route('/manager_dashboard')
def manager():
    if session.get('role') != 'Manager':
        return redirect('/login')
        
    current_users = read_sheet('SystemUsers')
    leaders = [u for u in current_users if u["role"] == "Team Leader"]
    projects = read_sheet('Projects')
    meetings = read_sheet('Meetings')
    
    available_files = os.listdir(PROJECT_FOLDER) if os.path.exists(PROJECT_FOLDER) else []

    return render_template("manager.html", 
                           users=current_users, 
                           leaders=leaders,
                           projects=projects, 
                           attendance=attendance, 
                           notifications=notifications, 
                           meetings=meetings,
                           available_files=available_files,
                           meeting_link=MEETING_LINK)

@app.route('/get_user_name')
def get_user_name():
    email = request.args.get('email', '').strip().lower()
    users = read_sheet('SystemUsers')
    for u in users:
        if str(u['email']).lower() == email:
            return jsonify({"status": "found", "name": u['name'], "role": u['role']})
    return jsonify({"status": "not_found", "name": "⚠️ Identity Record Not Found inside Database Layer"})

@app.route('/get_all_excel_users', methods=['GET'])
def get_all_excel_users():
    users = read_sheet('SystemUsers')
    return jsonify(users)

@app.route('/edit_file_source/<path:filename>')
@app.route('/view_workspace_code')
def edit_file_source(filename=None):
    if not filename:
        filename = request.args.get('filename', '')
        
    path = os.path.join(PROJECT_FOLDER, filename)
    if not os.path.abspath(path).startswith(os.path.abspath(PROJECT_FOLDER)):
        return "Access Scope Violation Exception", 403
        
    content = ""
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
    else:
        return f"File Matrix Node ({filename}) Not Found", 404
        
    return render_template("code_editor.html", filename=filename, content=content)

@app.route('/assign_project', methods=['POST'])
@app.route('/assign_project_leader', methods=['POST'])
def assign_project_leader():
    project_id = request.form.get('project_id', '').strip()
    leader_email = request.form.get('leader_email', '').strip()
    
    if not project_id:
        project_name = request.form.get('project_name', '').strip()
        projects = read_sheet('Projects')
        target_proj = next((p for p in projects if p["name"] == project_name), None)
        if target_proj:
            project_id = target_proj["id"]

    if not leader_email:
        leader_email = request.form.get('team_leader', '').strip()

    projects = read_sheet('Projects')
    current_users = read_sheet('SystemUsers')
    leader_name = next((u["name"] for u in current_users if u["email"].lower() == leader_email.lower()), leader_email)

    for p in projects:
        if str(p["id"]) == project_id or str(p["name"]) == request.form.get('project_name', '').strip():
            p["assigned_leader_email"] = leader_email
            if p["status"] in ["Pending", "Unassigned"]:
                p["status"] = "In Progress"
                p["progress"] = 15
            else:
                p["status"] = "Assigned"
            break

    write_sheet('Projects', projects)
    notifications.append(f"Project Code [{project_id}] explicitly delegated to: {leader_name}.")
    
    socketio.emit('live_notification', {
        "to": leader_email,
        "message": f"Administrative Management Framework has delegated access scope control of structural path: {project_id}"
    })
    return redirect("/manager_dashboard")

@app.route('/create_user', methods=['POST'])
def create_user():
    name = request.form.get('name', '').strip()
    email = request.form.get('email', '').strip()
    password = request.form.get('password', '1234').strip()
    role = request.form.get('role', 'Team Leader').strip()
    
    if name and email:
        users = read_sheet('SystemUsers')
        users.append({
            "name": name,
            "email": email,
            "password": password,
            "role": role
        })
        write_sheet('SystemUsers', users)
        notifications.append(f"Created secure user profile node for {name} ({role}).")
        
    return redirect("/manager_dashboard")

# =====================================================
# INTERACTIVE MEETINGS PIPELINE CONTROLLER
# =====================================================
@app.route('/schedule_meeting', methods=['POST'])
def schedule_meeting():
    title = request.form.get('title', 'System Matrix Sync').strip()
    host_email = session.get('user', 'manager@gmail.com')
    invitee_email = request.form.get('invitee_email', '').strip()
    scheduled_time = request.form.get('scheduled_time', '').strip()
    
    room_token = f"room_{random.randint(1000,9999)}"
    
    meetings = read_sheet('Meetings')
    meeting_data = {
        "meeting_id": f"M-{random.randint(100, 999)}",
        "title": title,
        "host_email": host_email,
        "invitee_email": invitee_email,
        "scheduled_time": scheduled_time,
        "room_token": room_token
    }
    meetings.append(meeting_data)
    write_sheet('Meetings', meetings)
    
    notifications.append(f"Meeting pipeline built successfully: {title} scheduled with {invitee_email}.")
    
    socketio.emit('live_notification', {
        "to": invitee_email,
        "message": f"Calendar Matrix Update: New conference interface requested by {host_email} at {scheduled_time}."
    })
    
    socketio.emit(
        'receive_meeting_push', 
        meeting_data, 
        to=f"dashboard_user_{invitee_email}"
    )
    
    if request.headers.get('X-Requested-With') == 'XMLHttpRequest' or request.is_json:
        return jsonify({"status": "success", "message": "Meeting generated and transmitted successfully."})
        
    if session.get('role') == 'Manager':
        return redirect('/manager_dashboard')
    elif session.get('role') == 'Team Leader':
        return redirect('/teamleader')
    return redirect('/login')

@app.route('/join_call_session/<room_token>')
def join_call_session(room_token):
    return f"<h1>WebRTC Channel Active: {room_token}</h1><p>Embedded video context active.</p>"

# =====================================================
# DYNAMIC TEAM LEADER DASHBOARD WORKSPACE
# =====================================================
@app.route('/teamleader')
@app.route('/leader')
def teamleader():
    email = request.args.get("email") or session.get("user") or "vaishusamalla11@gmail.com"
    session["user"] = email
    session["role"] = "Team Leader"
    
    all_projects = read_sheet('Projects')
    all_tasks = read_sheet('SubTasks')
    all_users = read_sheet('SystemUsers')
    meetings = read_sheet('Meetings')
    
    my_projects = [p for p in all_projects if str(p.get('assigned_leader_email')).lower() == email.lower()]
    my_tasks = [t for t in all_tasks if str(t.get('team_leader')).lower() == email.lower()]
    team_members = [u for u in all_users if u.get("role") == "Team Member"]
    my_meetings = [m for m in meetings if m.get('host_email') == email or m.get('invitee_email') == email]
    
    return render_template("teamleader.html", 
                           projects=my_projects, 
                           tasks=my_tasks, 
                           team_members=team_members,
                           meetings=my_meetings,
                           leader_email=email,
                           name=session.get('name', 'Team Leader'))

@app.route('/update_project_status', methods=['POST'])
def update_project_status():
    p_name = request.form.get('project_name', '').strip()
    proj_id = request.form.get('project_id', '').strip()
    new_status = request.form['status'].strip()
    new_progress = int(request.form['progress'].strip())

    projects = read_sheet('Projects')
    for p in projects:
        if (str(p["id"]) == proj_id or str(p["name"]) == p_name) and (str(p["assigned_leader_email"]).lower() == session.get('user', '').lower() or session.get('role') == 'Manager'):
            p["status"] = new_status
            p["progress"] = new_progress
            break

    write_sheet('Projects', projects)
    notifications.append(f"Project updated: {new_status} ({new_progress}%).")
    return redirect(url_for('teamleader'))

@app.route('/split_task', methods=['POST'])
def split_task():
    tasks = read_sheet('SubTasks')
    project_id = request.form.get('project', '').strip()
    module_type = request.form.get('module', '').strip()  
    assigned_member = request.form.get('assigned_to', '').strip()
    
    safe_type = "backend" if "backend" in module_type.lower() else "frontend"
    safe_filename = f"{project_id.replace('-', '')}_{safe_type}_workspace.txt"
    
    tasks.append({
        "project": project_id,
        "module": f"{module_type} Component Partition",
        "assigned_to": assigned_member,
        "status": "Assigned",
        "progress": 0,
        "team_leader": session.get('user', 'vaishusamalla11@gmail.com'),
        "filename": safe_filename
    })
    write_sheet('SubTasks', tasks)
    
    file_path = os.path.join(PROJECT_FOLDER, safe_filename)
    if not os.path.exists(file_path):
        with open(file_path, "w", encoding="utf-8") as f:
            f.write(f"// Cloud Workspace Pipeline initialized for {module_type}.\n// Begin drafting configuration block nodes.\n")
            
    socketio.emit('live_notification', {
        "to": assigned_member,
        "message": f"New functional module tracking sequence allocated: '{module_type}' under structural layout envelope."
    })
    return redirect(url_for('teamleader'))

# =====================================================
# TEAM MEMBER INTERACTION INTERFACES
# =====================================================
@app.route('/teammember')
@app.route('/member')
def teammember():
    email = request.args.get("email") or session.get("user") or "emulaabhi@gmail.com"
    session["user"] = email
    session["role"] = "Team Member"
    
    all_tasks = read_sheet('SubTasks')
    meetings = read_sheet('Meetings')
    
    member_tasks = [t for t in all_tasks if str(t.get('assigned_to')).lower() == email.lower()]
    my_meetings = [m for m in meetings if m.get('invitee_email') == email or m.get('host_email') == email]
    
    return render_template("teammember.html", tasks=member_tasks, meetings=my_meetings, member_email=email)

@app.route('/update_task/<int:task_index>', methods=['POST'])
def update_task(task_index):
    user_email = session.get('email') or session.get('user')
    all_tasks = read_sheet('SubTasks')
    
    counter = 0
    for task in all_tasks:
        if str(task.get('assigned_to')).lower() == user_email.lower():
            if counter == task_index:
                task['status'] = request.form.get('status')
                task['progress'] = int(request.form.get('progress'))
                break
            counter += 1
            
    write_sheet('SubTasks', all_tasks)
    return redirect(url_for('teammember'))

# =====================================================
# CLIENT INTERFACES
# =====================================================
@app.route('/client')
def client():
    projects = read_sheet('Projects')
    meetings = read_sheet('Meetings')
    user_email = session.get('user', 'client@gmail.com')
    
    my_meetings = [m for m in meetings if m.get('invitee_email') == user_email or m.get('host_email') == user_email]
    return render_template("client.html", projects=projects, meetings=my_meetings, meeting_link=MEETING_LINK)

@app.route('/submit_project', methods=['POST'])
def submit_project():
    projects = read_sheet('Projects')
    new_id = f"PROJ-{random.randint(100, 999)}"
    projects.append({
        "id": new_id,
        "name": request.form['project'].strip(),
        "description": request.form['description'].strip(),
        "deadline": request.form['deadline'],
        "status": "Pending",
        "assigned_leader_email": "",
        "progress": 0
    })
    write_sheet('Projects', projects)
    return redirect("/client")

# =====================================================
# COLLABORATIVE UTILITIES & SOURCE ARCHIVES
# =====================================================
@app.route('/save_code', methods=['POST'])
def save_code():
    data = request.get_json() or {}
    filename = data.get('filename')
    code_content = data.get('code', '')
    if not filename:
        return jsonify({"error": "No script identity mapped"}), 400
    path = os.path.join(PROJECT_FOLDER, filename)
    with open(path, "w", encoding="utf-8") as f:
        f.write(code_content)
    return jsonify({"success": True, "message": "Written down to disk safely."})

@app.route('/load_code/<filename>')
def load_code(filename):
    path = os.path.join(PROJECT_FOLDER, filename)
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return jsonify({"code": f.read()})
    return jsonify({"code": "// Project storage record initialization node sequence placeholder empty block."})

@app.route('/logout')
def logout():
    session.clear()
    return redirect('/login')

# =====================================================
# WEBRTC MULTI-STREAM MEDIA ROUTING ENGINE
# =====================================================
@socketio.on('register_dashboard_identity')
def on_register_identity(data):
    user_email = data.get('email')
    if user_email:
        personal_room = f"dashboard_user_{user_email}"
        join_room(personal_room)
        active_user_dashboards[user_email] = request.sid
        emit('status_message', {'message': f'Dashboard data pipeline locked onto node: {user_email}'})

@socketio.on('join_room')
def handle_join_room_event(data):
    username = data.get('username', 'Anonymous User')
    room = data.get('room', 'public_lobby')
    
    join_room(room)
    payload = {
        'msg': f"🔒 Security Notice: {username} has connected to communication stream pipeline: {room}.",
        'message': f"✨ {username} connected inside call target zone: {room.replace('_', ' ').title()}"
    }
    emit('status_message', payload, to=room)

@socketio.on('send_chat_message')
def handle_chat_delivery(data):
    room = data.get('room', 'public_lobby')
    sender = data.get('sender', 'System Admin')
    message = data.get('message') or data.get('msg')
    timestamp = datetime.now().strftime("%H:%M:%S")
    
    payload = {
        'room': room,
        'sender': sender,
        'msg': message,
        'message': message,
        'timestamp': timestamp
    }
    emit('receive_message', payload, to=room)

@socketio.on('toggle_media_track')
def handle_media_stream_signals(data):
    room = data.get('room', 'public_lobby')
    media_type = data.get('type')  
    is_enabled = data.get('enabled')
    sender = data.get('sender', 'Anonymous')
    
    payload = {
        'user': sender,
        'sender': sender,
        'type': media_type,
        'enabled': is_enabled
    }
    emit('media_state_altered', payload, to=room)

@socketio.on('webrtc_signal')
def handle_webrtc_pass_through(data):
    room = data.get('room')
    emit('webrtc_response', data, to=room, include_self=False)

# Changed execution loop to catch Render's system port assignments seamlessly
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5050))
    print(f"🚀 Unified Flowboard Media Engine Infrastructure Active on Port: {port}")
    socketio.run(app, host="0.0.0.0", port=port, use_reloader=False)