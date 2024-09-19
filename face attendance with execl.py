import face_recognition
import os
import cv2
import pandas as pd
from datetime import datetime

# Function to load roll numbers and names from a text file
def load_roll_numbers_and_names(file_path):
    roll_number_map = {}
    with open(file_path, 'r') as file:
        for line in file:
            roll_number, name = line.strip().split(',', 1)
            roll_number_map[name.strip()] = roll_number.strip()
    return roll_number_map

# Function to register a new person
def register_person(name, folder_path):
    print(f"Registering {name}...")
    images = []

    # Capture multiple images for better face recognition accuracy
    for i in range(5):
        input("Press Enter to capture an image...")
        _, frame = cap.read()
        face_locations = face_recognition.face_locations(frame)
        face_encodings = face_recognition.face_encodings(frame, face_locations)

        if len(face_encodings) > 0:
            images.append(face_encodings[0])

    # Save the collected images and name for future recognition
    if images:
        for i, encoding in enumerate(images):
            file_path = os.path.join(folder_path, f"{name}_{i}.jpg")
            cv2.imwrite(file_path, frame)

# Function to get the list of registered images and corresponding names from the specified folder
def get_registered_images_and_names(folder_path):
    images = []
    names = []

    for filename in os.listdir(folder_path):
        if filename.lower().endswith(('.png', '.jpg', '.jpeg')):
            image_path = os.path.join(folder_path, filename)
            image = face_recognition.load_image_file(image_path)
            encoding = face_recognition.face_encodings(image)

            if len(encoding) > 0:
                images.append(encoding[0])
                names.append(os.path.splitext(filename)[0].split('_')[0])  # Extract name from filename

    return images, names

# Function to mark attendance in an Excel file
def mark_attendance(name, excel_file_path, roll_number_map):
    now = datetime.now()
    date_string = now.strftime("%Y-%m-%d")
    time_string = now.strftime("%H:%M:%S")
    day_string = now.strftime("%A")
    roll_number = roll_number_map.get(name, "N/A")

    # Load the existing data if the file exists, otherwise create a new DataFrame
    if os.path.exists(excel_file_path):
        df = pd.read_excel(excel_file_path)
    else:
        df = pd.DataFrame(columns=["Roll Number", "Name", "Date", "Time", "Day"])

    # Create a new DataFrame for the new record
    new_record = pd.DataFrame([{"Roll Number": roll_number, "Name": name, "Date": date_string, "Time": time_string, "Day": day_string}])
    
    # Concatenate the existing DataFrame with the new record
    df = pd.concat([df, new_record], ignore_index=True)

    # Save the DataFrame back to the Excel file
    df.to_excel(excel_file_path, index=False)
    print(f"{name}'s attendance marked.")

# Get folder path, Excel file path, text file path for roll numbers, and registration flag from the user
folder_path = input("Enter the path of the folder for registration and recognition: ")
excel_file_path = input("Enter the path of the Excel file for attendance: ")
roll_number_file_path = input("Enter the path of the text file with roll numbers and names: ")
register_flag = input("Do you want to register new persons? (y/n): ").lower() == 'y'

# Load roll numbers and names
roll_number_map = load_roll_numbers_and_names(roll_number_file_path)

# Get the list of registered images and names
if register_flag:
    # Registration step
    name = input("Enter the name of the person to register: ")
    register_person(name, folder_path)

# Get the list of registered images and names
known_face_encodings, known_face_names = get_registered_images_and_names(folder_path)

# Initialize the webcam
cap = cv2.VideoCapture(0)

# Set to keep track of marked names
marked_names = set()

while True:
    ret, frame = cap.read()

    # Find all face locations and face encodings in the current frame
    face_locations = face_recognition.face_locations(frame)
    face_encodings = face_recognition.face_encodings(frame, face_locations)

    for (top, right, bottom, left), face_encoding in zip(face_locations, face_encodings):
        # Check if the face matches any registered faces
        matches = face_recognition.compare_faces(known_face_encodings, face_encoding)
        name = "Unknown"

        if True in matches:
            first_match_index = matches.index(True)
            name = known_face_names[first_match_index]
            if name not in marked_names:
                mark_attendance(name, excel_file_path, roll_number_map)  # Mark attendance for the recognized person
                marked_names.add(name)

        # Draw a rectangle around the face and display the name
        cv2.rectangle(frame, (left, top), (right, bottom), (0, 255, 0), 2)
        font = cv2.FONT_HERSHEY_DUPLEX
        cv2.putText(frame, name, (left + 6, bottom - 6), font, 0.5, (255, 255, 255), 1)

    # Display the resulting frame
    cv2.imshow('Face Recognition Attendance System', frame)

    # Exit condition
    if cv2.waitKey(1) == ord('q'):
        break

# Release the webcam and close all windows
cap.release()
cv2.destroyAllWindows()
