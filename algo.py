import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment
import logging
import os
from datetime import datetime
from collections import defaultdict
import math

# Configure logging
logging.basicConfig(
    filename='errors.txt',
    level=logging.ERROR,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

class ExamSeatingArrangement:
    def __init__(self, buffer=1, mode='dense'):
        """
        Initialize the exam seating arrangement system.
        
        Args:
            buffer (int): Number of buffer seats per classroom
            mode (str): 'dense' or 'sparse' seating mode
        """
        self.buffer = buffer
        self.mode = mode
        self.validate_mode()
        
        # Data storage
        self.exam_schedule = None
        self.course_roll_mapping = None
        self.roll_name_mapping = None
        self.classroom_master = None
        
        # Output containers
        self.seating_arrangement = []
        self.seats_left = []
        self.clashes = []
    
    def validate_mode(self):
        """Validate the seating mode is either 'dense' or 'sparse'."""
        if self.mode not in ['dense', 'sparse']:
            raise ValueError("Mode must be either 'dense' or 'sparse'")
    
    def load_data(self, file_path):
        """
        Load data from the input Excel file with multiple sheets.
        
        Args:
            file_path (str): Path to the input Excel file
        """
        try:
            xls = pd.ExcelFile(file_path)
            
            # Load each sheet
            self.exam_schedule = pd.read_excel(xls, sheet_name='in_timetable')
            self.course_roll_mapping = pd.read_excel(xls, sheet_name='in_course_roll_mapping')
            self.roll_name_mapping = pd.read_excel(xls, sheet_name='in_roll_name_mapping')
            self.classroom_master = pd.read_excel(xls, sheet_name='in_room_capacity')
            
            # Clean and standardize column names
            self.standardize_data()
            
        except Exception as e:
            logging.error(f"Error loading data from {file_path}: {str(e)}")
            raise
    
    def standardize_data(self):
        """Standardize column names and data formats across dataframes."""
        try:
            # Standardize column names
            self.exam_schedule.columns = [col.strip().lower() for col in self.exam_schedule.columns]
            self.course_roll_mapping.columns = [col.strip().lower() for col in self.course_roll_mapping.columns]
            self.roll_name_mapping.columns = [col.strip().lower() for col in self.roll_name_mapping.columns]
            self.classroom_master.columns = [col.strip().lower() for col in self.classroom_master.columns]

            # Parse and expand the exam schedule
            self.exam_schedule = self.parse_exam_schedule(self.exam_schedule)

            # Ensure roll_no is string type
            if 'rollno' in self.course_roll_mapping.columns:
                self.course_roll_mapping['rollno'] = self.course_roll_mapping['rollno'].astype(str)
            if 'roll' in self.roll_name_mapping.columns:
                self.roll_name_mapping['roll'] = self.roll_name_mapping['roll'].astype(str)

            # Fill missing names with "Unknown Name"
            if 'name' in self.roll_name_mapping.columns:
                self.roll_name_mapping['name'] = self.roll_name_mapping['name'].fillna('Unknown Name')
            
        except Exception as e:
            logging.error(f"Error standardizing data: {str(e)}")
            raise
    
    def parse_exam_schedule(self, schedule_df):
        """Parse the exam schedule with combined subject codes."""
        parsed_rows = []
        
        for _, row in schedule_df.iterrows():
            date = row['date']
            day = row['day']
            
            # Process morning session
            if pd.notna(row['morning']) and str(row['morning']).strip().lower() != 'no exam':
                subjects = [s.strip() for s in str(row['morning']).split(';') if s.strip()]
                for subject in subjects:
                    parsed_rows.append({
                        'date': date,
                        'day': day,
                        'session': 'morning',
                        'subject_code': subject
                    })
            
            # Process evening session
            if pd.notna(row['evening']) and str(row['evening']).strip().lower() != 'no exam':
                subjects = [s.strip() for s in str(row['evening']).split(';') if s.strip()]
                for subject in subjects:
                    parsed_rows.append({
                        'date': date,
                        'day': day,
                        'session': 'evening',
                        'subject_code': subject
                    })
        
        return pd.DataFrame(parsed_rows)
    
    def detect_clashes(self):
        """Detect exam clashes for students."""
        try:
            # Group exams by date and session
            grouped_exams = self.exam_schedule.groupby(['date', 'day', 'session'])
            
            for (date, day, session), exams in grouped_exams:
                subject_codes = exams['subject_code'].unique()
                
                # Get all unique pairs of subjects
                from itertools import combinations
                for subj1, subj2 in combinations(subject_codes, 2):
                    # Get roll numbers for each subject
                    rolls1 = set(self.get_rolls_for_subject(subj1))
                    rolls2 = set(self.get_rolls_for_subject(subj2))
                    
                    # Find intersection
                    common_rolls = rolls1 & rolls2
                    
                    if common_rolls:
                        clash_info = {
                            'date': date,
                            'day': day,
                            'session': session,
                            'subjects': [subj1, subj2],
                            'roll_numbers': list(common_rolls)
                        }
                        self.clashes.append(clash_info)
                        
                        # Log the clash
                        for roll in common_rolls:
                            logging.error(
                                f"Roll No {roll} has multiple exams on {date}, {session} in {subj1} and {subj2}"
                            )
                            print(
                                f"Roll No {roll} has multiple exams on {date}, {session} in {subj1} and {subj2}"
                            )
            
            return len(self.clashes) == 0  # Returns True if no clashes
        
        except Exception as e:
            logging.error(f"Error detecting clashes: {str(e)}")
            raise
    
    def get_rolls_for_subject(self, subject_code):
        """Get all roll numbers for a given subject code."""
        try:
            rolls = self.course_roll_mapping[
                self.course_roll_mapping['course_code'] == subject_code
            ]['rollno'].unique()
            return list(rolls)
        except Exception as e:
            logging.error(f"Error getting rolls for subject {subject_code}: {str(e)}")
            return []
    
    def get_student_name(self, roll_no):
        """Get student name for a given roll number."""
        try:
            name = self.roll_name_mapping[
                self.roll_name_mapping['roll'] == str(roll_no)
            ]['name'].values[0]
            return name
        except (IndexError, KeyError):
            logging.warning(f"Name not found for roll number {roll_no}")
            return "Unknown Name"
    
    def allocate_seats(self):
        """Allocate seats for all exams based on the given constraints."""
        try:
            # Group exams by date and session
            grouped_exams = self.exam_schedule.groupby(['date', 'day', 'session'])
            
            for (date, day, session), exams in grouped_exams:
                subject_codes = exams['subject_code'].unique()
                
                # Sort subjects by number of students (descending)
                subjects_with_counts = []
                for subj in subject_codes:
                    rolls = self.get_rolls_for_subject(subj)
                    subjects_with_counts.append((subj, len(rolls)))
                
                # Sort by student count descending
                subjects_with_counts.sort(key=lambda x: x[1], reverse=True)
                
                # Prepare available rooms (sorted by capacity descending)
                available_rooms = self.classroom_master.sort_values(
                    'exam capacity', ascending=False
                ).copy()
                
                # Apply buffer to room capacities
                available_rooms['adjusted_capacity'] = available_rooms['exam capacity'] - self.buffer
                
                # Allocate rooms for each subject
                for subj, student_count in subjects_with_counts:
                    self.allocate_subject_rooms(
                        subj, student_count, available_rooms, date, day, session
                    )
            
            # Prepare seats left data
            self.prepare_seats_left_data()
            
        except Exception as e:
            logging.error(f"Error allocating seats: {str(e)}")
            raise

    def select_best_room(self, available_rooms, remaining_students):
        """Select the most appropriate room based on mode and capacity."""
        try:
            if self.mode == 'dense':
                # Find smallest room that can fit all remaining students
                suitable_rooms = available_rooms[
                    available_rooms['adjusted_capacity'] >= remaining_students
                ]
                if not suitable_rooms.empty:
                    return suitable_rooms.iloc[-1]  # Take the smallest suitable room
                else:
                    return available_rooms.iloc[0]  # Take the largest available room
            else:  # sparse mode
                return available_rooms.iloc[0]  # Take the largest available room
        except Exception as e:
            logging.error(f"Error selecting best room: {str(e)}")
            return None
    
    def allocate_subject_rooms(self, subject_code, student_count, available_rooms, date, day, session):
        """Allocate rooms for a specific subject with proper capacity enforcement."""
        try:
            remaining_students = student_count
            rolls = self.get_rolls_for_subject(subject_code)
            
            while remaining_students > 0 and not available_rooms.empty:
                # Select best room based on mode
                room = self.select_best_room(available_rooms, remaining_students)
                if room is None:
                    break
                    
                # Calculate how many students we can allocate (ACTUAL count)
                students_to_allocate = min(
                    remaining_students,
                    math.floor(room['adjusted_capacity'] * 0.5) if self.mode == 'sparse' 
                    else room['adjusted_capacity']
                )
                
                # Get ACTUAL roll numbers to allocate
                allocated_rolls = rolls[:students_to_allocate]
                actual_count = len(allocated_rolls)  # ACTUAL number of students
                
                # Record allocation with ACTUAL count
                self.seating_arrangement.append({
                    'date': date,
                    'day': day,
                    'session': session,
                    'course_code': subject_code,
                    'room': room['room no.'],
                    'block': room['block'],
                    'allocated_student_count': actual_count,  # Store ACTUAL count
                    'roll_list': ';'.join(allocated_rolls),
                    'exam_capacity': room['exam capacity'],
                    'adjusted_capacity': room['adjusted_capacity']
                })
                
                # Update remaining students and room availability
                remaining_students -= actual_count
                rolls = rolls[actual_count:]  # Remove allocated students
                
                # Update room capacity
                available_rooms.at[room.name, 'adjusted_capacity'] -= actual_count
                if available_rooms.at[room.name, 'adjusted_capacity'] <= 0:
                    available_rooms = available_rooms.drop(room.name)
                else:
                    available_rooms = available_rooms.sort_values('adjusted_capacity', ascending=False)
                
            if remaining_students > 0:
                error_msg = f"Cannot allocate all students for {subject_code} on {date} {session}. {remaining_students} students left."
                logging.error(error_msg)
                print(error_msg)
                
        except Exception as e:
            logging.error(f"Error allocating rooms for {subject_code}: {str(e)}")
            raise

    def prepare_seats_left_data(self):
        """Calculate remaining capacity with proper validation."""
        try:
            if not self.seating_arrangement:
                return
                
            # Create DataFrame with ACTUAL student counts
            room_usage = pd.DataFrame(self.seating_arrangement)
            room_usage['actual_count'] = room_usage['roll_list'].apply(
                lambda x: len(x.split(';')) if pd.notna(x) else 0
            )
            
            # Group by room and calculate totals
            room_allocation = room_usage.groupby(['room', 'block', 'exam_capacity']).agg({
                'actual_count': 'sum'
            }).reset_index()
            
            # Calculate vacant seats properly
            room_allocation['vacant'] = room_allocation['exam_capacity'] - room_allocation['actual_count']
            room_allocation['vacant'] = room_allocation['vacant'].clip(lower=0)  # No negative vacancies
            
            # Convert to output format
            self.seats_left = room_allocation[[
                'room', 'exam_capacity', 'block', 'actual_count', 'vacant'
            ]].rename(columns={
                'room': 'room_no',
                'actual_count': 'allocated'
            }).to_dict('records')
            
        except Exception as e:
            logging.error(f"Error preparing seats left data: {str(e)}")
            raise
    
    def generate_output_files(self, output_folder='output'):
        """Generate all output files and folder structure."""
        try:
            # Create main output folder
            os.makedirs(output_folder, exist_ok=True)
            
            # Generate overall files
            self.generate_overall_files(output_folder)
            
            # Generate per-date files
            self.generate_per_date_files(output_folder)
            
        except Exception as e:
            logging.error(f"Error generating output files: {str(e)}")
            raise
    
    def generate_overall_files(self, output_folder):
        """Generate the overall seating arrangement and seats left files."""
        try:
            # Overall seating arrangement
            overall_seating = pd.DataFrame(self.seating_arrangement)
            overall_seating = overall_seating[[
                'date', 'day', 'course_code', 'room', 'allocated_student_count', 'roll_list'
            ]]
            
            # Save to Excel with formatting
            output_path = os.path.join(output_folder, 'op_overall_seating_arrangement.xlsx')
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                overall_seating.to_excel(writer, index=False, sheet_name='Seating Plan')
                
                # Format the worksheet
                workbook = writer.book
                worksheet = writer.sheets['Seating Plan']
                
                # Add title
                worksheet.insert_rows(1)
                worksheet.cell(row=1, column=1, value="Seating Plan")
                worksheet.cell(row=1, column=1).font = Font(bold=True, size=14)
                worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
            
            # Seats left file
            seats_left_df = pd.DataFrame(self.seats_left)
            seats_left_df.to_excel(
                os.path.join(output_folder, 'op_seats_left.xlsx'),
                index=False
            )
            
        except Exception as e:
            logging.error(f"Error generating overall files: {str(e)}")
            raise
    
    def generate_per_date_files(self, output_folder):
        """Generate the per-date folder structure and files."""
        try:
            # Group seating arrangement by date
            seating_df = pd.DataFrame(self.seating_arrangement)
            if seating_df.empty:
                return
                
            grouped = seating_df.groupby(['date', 'session'])
            
            for (date, session), group in grouped:
                # Format date for folder name
                date_str = pd.to_datetime(date).strftime('%d_%m_%Y')
                session_str = 'morning' if session == 'morning' else 'evening'
                
                # Create date and session folders
                date_folder = os.path.join(output_folder, date_str)
                session_folder = os.path.join(date_folder, session_str)
                os.makedirs(session_folder, exist_ok=True)
                
                # Group by subject and room
                subject_room_group = group.groupby(['course_code', 'room'])
                
                for (subject_code, room), subject_group in subject_room_group:
                    # Generate attendance sheet
                    self.generate_attendance_sheet(
                        subject_code, room, subject_group, session_folder
                    )
                    
                    # Generate seating arrangement file
                    self.generate_seating_arrangement_file(
                        subject_code, room, subject_group, session_folder
                    )
                    
        except Exception as e:
            logging.error(f"Error generating per-date files: {str(e)}")
            raise
    
    def generate_attendance_sheet(self, subject_code, room, subject_group, session_folder):
        """Generate attendance sheet for a subject in a specific room."""
        try:
            # Get roll numbers and names
            roll_list = subject_group['roll_list'].values[0].split(';')
            student_data = []
            
            for roll in roll_list:
                student_data.append({
                    'Roll No': roll,
                    'Student Name': self.get_student_name(roll),
                    'Signature': ''
                })
            
            # Create DataFrame
            attendance_df = pd.DataFrame(student_data)
            
            # Save to Excel
            filename = f"attendance_sheet_{subject_code}_{room}.xlsx"
            filepath = os.path.join(session_folder, filename)
            attendance_df.to_excel(filepath, index=False)
            
        except Exception as e:
            logging.error(f"Error generating attendance sheet for {subject_code} in {room}: {str(e)}")
            raise
    
    def generate_seating_arrangement_file(self, subject_code, room, subject_group, session_folder):
        """Generate seating arrangement file for a subject in a specific room."""
        try:
            # Get the first row (all rows should have same date, day, etc.)
            row = subject_group.iloc[0]
            
            # Create DataFrame with the arrangement info
            arrangement_data = [{
                'date': row['date'],
                'day': row['day'],
                'course_code': row['course_code'],
                'room': row['room'],
                'allocated_student_count': row['allocated_student_count'],
                'roll_list': row['roll_list']
            }]
            arrangement_df = pd.DataFrame(arrangement_data)
            
            # Save to Excel with formatting
            filename = f"seating_arrangement_{subject_code}_{room}.xlsx"
            filepath = os.path.join(session_folder, filename)
            
            with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                arrangement_df.to_excel(writer, index=False, sheet_name='Seating Plan')
                
                # Format the worksheet
                workbook = writer.book
                worksheet = writer.sheets['Seating Plan']
                
                # Add title
                worksheet.insert_rows(1)
                worksheet.cell(row=1, column=1, value="Seating Plan")
                worksheet.cell(row=1, column=1).font = Font(bold=True, size=14)
                worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
                
                # Add TA and invigilator sections
                start_row = len(arrangement_df) + 4
                for i in range(1, 6):
                    worksheet.cell(row=start_row + i - 1, column=1, value=f"TA{i}:")
                
                for i in range(1, 6):
                    worksheet.cell(row=start_row + 5 + i - 1, column=1, value=f"Invigilator {i}:")
            
        except Exception as e:
            logging.error(f"Error generating seating arrangement for {subject_code} in {room}: {str(e)}")
            raise


def main():
    try:
        # User inputs
        buffer_seats = int(input("Enter number of buffer seats per classroom: "))
        mode = input("Enter seating mode (dense/sparse): ").lower()
        
        # Initialize system
        seating_system = ExamSeatingArrangement(buffer=buffer_seats, mode=mode)
        
        # Load data
        input_file = 'sorted_output_file.xlsx'
        seating_system.load_data(input_file)
        
        # Detect clashes
        print("\nChecking for exam clashes...")
        no_clashes = seating_system.detect_clashes()
        
        if no_clashes:
            print("No exam clashes detected.")
        else:
            print("Exam clashes detected. Check errors.txt for details.")
        
        # Allocate seats
        print("\nAllocating seats...")
        seating_system.allocate_seats()
        
        # Generate output files
        print("\nGenerating output files...")
        seating_system.generate_output_files()
        
        print("\nProcess completed successfully!")
        
    except Exception as e:
        logging.error(f"Error in main execution: {str(e)}")
        print(f"An error occurred: {str(e)}. Check errors.txt for details.")


if __name__ == "__main__":
    main()