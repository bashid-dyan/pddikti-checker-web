"""
PDDIKTI Checker — Flask Web App
"""

import os
import uuid
import json
import threading
from datetime import datetime
from flask import Flask, request, jsonify, send_file, Response, render_template
from werkzeug.utils import secure_filename
from checker import run_checker

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, 'uploads')
RESULT_DIR = os.path.join(BASE_DIR, 'results')
HISTORY_FILE = os.path.join(BASE_DIR, 'history.json')

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(RESULT_DIR, exist_ok=True)

# In-memory job tracking
jobs = {}
# Lock for thread-safe history file access
history_lock = threading.Lock()


def load_history():
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return []


def save_history(history):
    with history_lock:
        with open(HISTORY_FILE, 'w', encoding='utf-8') as f:
            json.dump(history, f, ensure_ascii=False, indent=2)


def process_job(job_id, input_path, output_path, original_filename):
    """Background worker for processing a job."""
    try:
        jobs[job_id]['status'] = 'processing'

        def on_progress(current, total, nama, status_text):
            jobs[job_id]['current'] = current
            jobs[job_id]['total'] = total
            jobs[job_id]['logs'].append({
                'index': current,
                'total': total,
                'nama': nama,
                'status': status_text
            })
            # Keep only last 200 log entries in memory
            if len(jobs[job_id]['logs']) > 200:
                jobs[job_id]['logs'] = jobs[job_id]['logs'][-200:]

        summary = run_checker(input_path, output_path, on_progress=on_progress)

        jobs[job_id]['status'] = 'done'
        jobs[job_id]['summary'] = summary

        # Save to history
        history = load_history()
        history.insert(0, {
            'job_id': job_id,
            'filename': original_filename,
            'date': datetime.now().strftime('%Y-%m-%d %H:%M'),
            'total': summary['total'],
            'found': summary['found'],
            'not_found': summary['not_found'],
            'no_match': summary['no_match'],
        })
        # Keep last 50 entries
        history = history[:50]
        save_history(history)

    except Exception as e:
        jobs[job_id]['status'] = 'error'
        jobs[job_id]['error'] = str(e)


# ============================================================
# ROUTES
# ============================================================

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return jsonify({'error': 'Tidak ada file'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'Tidak ada file dipilih'}), 400

    if not file.filename.lower().endswith('.xlsx'):
        return jsonify({'error': 'File harus berformat .xlsx'}), 400

    job_id = str(uuid.uuid4())[:8]
    original_filename = secure_filename(file.filename)

    input_path = os.path.join(UPLOAD_DIR, f"{job_id}_{original_filename}")
    output_path = os.path.join(RESULT_DIR, f"{job_id}_Hasil_{original_filename}")

    file.save(input_path)

    jobs[job_id] = {
        'status': 'queued',
        'current': 0,
        'total': 0,
        'logs': [],
        'summary': None,
        'error': None,
        'filename': original_filename,
        'output_path': output_path,
        'created': datetime.now().strftime('%Y-%m-%d %H:%M'),
    }

    thread = threading.Thread(
        target=process_job,
        args=(job_id, input_path, output_path, original_filename),
        daemon=True
    )
    thread.start()

    return jsonify({'job_id': job_id})


@app.route('/progress/<job_id>')
def progress(job_id):
    """SSE endpoint for realtime progress."""
    def generate():
        last_sent = 0
        while True:
            job = jobs.get(job_id)
            if not job:
                yield f"data: {json.dumps({'error': 'Job tidak ditemukan'})}\n\n"
                break

            logs_to_send = job['logs'][last_sent:]
            last_sent = len(job['logs'])

            payload = {
                'status': job['status'],
                'current': job['current'],
                'total': job['total'],
                'logs': logs_to_send,
            }

            if job['status'] == 'done':
                payload['summary'] = job['summary']
                yield f"data: {json.dumps(payload, ensure_ascii=False)}\n\n"
                break
            elif job['status'] == 'error':
                payload['error'] = job['error']
                yield f"data: {json.dumps(payload, ensure_ascii=False)}\n\n"
                break

            yield f"data: {json.dumps(payload, ensure_ascii=False)}\n\n"

            import time
            time.sleep(1)

    return Response(generate(), mimetype='text/event-stream',
                    headers={'Cache-Control': 'no-cache', 'X-Accel-Buffering': 'no'})


@app.route('/download/<job_id>')
def download(job_id):
    job = jobs.get(job_id)

    # Check in-memory jobs first
    if job and job['status'] == 'done' and os.path.exists(job['output_path']):
        return send_file(
            job['output_path'],
            as_attachment=True,
            download_name=f"Hasil_{job['filename']}"
        )

    # Fallback: check results directory for historical downloads
    for f in os.listdir(RESULT_DIR):
        if f.startswith(job_id):
            return send_file(
                os.path.join(RESULT_DIR, f),
                as_attachment=True,
                download_name=f
            )

    return jsonify({'error': 'File tidak ditemukan'}), 404


@app.route('/history')
def history():
    return jsonify(load_history())


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
