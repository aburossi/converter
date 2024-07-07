"""This Python script converts XML data to JSON format through a graphical user interface (GUI) using Tkinter.
Works only with MC Questions. Export ZIP file from OLAT and select the root folder containing the XML files.
**Improve**: get the logic for KPRIM, Truefalse, and other types of questions."""

import os
import xml.etree.ElementTree as ET
import json
import tkinter as tk
from tkinter import filedialog

def parse_xml_file(file_path):
    tree = ET.parse(file_path)
    root = tree.getroot()

    # Extract question title
    title = root.get('title')

    # Extract question text
    question_text = root.find('.//{http://www.imsglobal.org/xsd/imsqti_v2p1}itemBody/{http://www.imsglobal.org/xsd/imsqti_v2p1}p').text

    # Extract answers
    answers = []
    correct_response = root.find('.//{http://www.imsglobal.org/xsd/imsqti_v2p1}correctResponse/{http://www.imsglobal.org/xsd/imsqti_v2p1}value').text
    choices = root.findall('.//{http://www.imsglobal.org/xsd/imsqti_v2p1}simpleChoice')
    
    for choice in choices:
        answer = {
            "text": choice.find('{http://www.imsglobal.org/xsd/imsqti_v2p1}p').text,
            "is_correct": choice.get('identifier') == correct_response,
            "feedback": ""  # Initialize feedback as empty string
        }
        answers.append(answer)

    # Extract feedback
    feedbacks = root.findall('.//{http://www.imsglobal.org/xsd/imsqti_v2p1}modalFeedback')
    for feedback in feedbacks:
        identifier = feedback.get('identifier')
        feedback_text = feedback.find('{http://www.imsglobal.org/xsd/imsqti_v2p1}p').text
        for answer in answers:
            if answer['is_correct'] and identifier.endswith('735'):  # Assuming correct feedback ends with 735
                answer['feedback'] = feedback_text
            elif not answer['is_correct'] and identifier.endswith(('745', '760')):  # Assuming incorrect feedbacks end with 745 or 760
                answer['feedback'] = feedback_text

    # Extract points
    max_score = float(root.find('.//{http://www.imsglobal.org/xsd/imsqti_v2p1}outcomeDeclaration[@identifier="MAXSCORE"]/{http://www.imsglobal.org/xsd/imsqti_v2p1}defaultValue/{http://www.imsglobal.org/xsd/imsqti_v2p1}value').text)

    return {
        "title": title,
        "type": "MC",
        "question_text": question_text,
        "answers": answers,
        "points": max_score,
        "penalty": 0,
        "additional_info": {}
    }
def process_folders(root_dir):
    questions = []
    for folder in os.listdir(root_dir):
        folder_path = os.path.join(root_dir, folder)
        if os.path.isdir(folder_path):
            for file in os.listdir(folder_path):
                if file.endswith('.xml'):
                    file_path = os.path.join(folder_path, file)
                    question = parse_xml_file(file_path)
                    questions.append(question)
    return questions

def main():
    # Create a Tkinter root window
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    # Ask the user for the root folder
    root_dir = filedialog.askdirectory(title="Select the root folder")

    # Process the folders and parse the XML files
    questions = process_folders(root_dir)
    
    # Ask the user for the location to save the JSON file
    save_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")], title="Save JSON file as")

    # Save the questions to a JSON file
    json_output = {"questions": questions}
    
    with open(save_path, 'w', encoding='utf-8') as f:
        json.dump(json_output, f, ensure_ascii=False, indent=2)

if __name__ == "__main__":
    main()
