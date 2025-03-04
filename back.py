from flask import Flask, request, jsonify, send_file
from flask_sqlalchemy import SQLAlchemy
import pandas as pd
import os

app = Flask(__name__)

# Use PostgreSQL (Replace <username>, <password>, <dbname>)
app.config['SQLALCHEMY_DATABASE_URI'] = 'postgresql://<username>:<password>@localhost:5432/<dbname>'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)

# Define the Database Model
class TimeMotionData(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    task_name = db.Column(db.String(255), nullable=False)
    start_time = db.Column(db.String(255), nullable=False)
    end_time = db.Column(db.String(255), nullable=False)
    duration = db.Column(db.String(255), nullable=False)

# Create the database tables
with app.app_context():
    db.create_all()

@app.route('/submit', methods=['POST'])
def submit_data():
    data = request.json
    new_entry = TimeMotionData(
        task_name=data['task_name'],
        start_time=data['start_time'],
        end_time=data['end_time'],
        duration=data['duration']
    )
    db.session.add(new_entry)
    db.session.commit()
    return jsonify({"message": "Data saved successfully!"})

@app.route('/export', methods=['GET'])
def export_excel():
    data = TimeMotionData.query.all()
    data_dict = [{"Task": d.task_name, "Start": d.start_time, "End": d.end_time, "Duration": d.duration} for d in data]
    df = pd.DataFrame(data_dict)

    file_path = "time_motion_data.xlsx"
    df.to_excel(file_path, index=False)

    return send_file(file_path, as_attachment=True)

@app.route('/')
def home():
    return "Time-Motion Study App is Running!"

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
