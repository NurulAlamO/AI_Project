import pandas as pd
import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
TOTAL_CLASSES = 20   
INPUT_FILE_NAME = 'student_data.xlsx'
OUTPUT_FILE_NAME = 'Attendance_Report.xlsx'

def process_attendance_from_excel():
    try:
        df_input = pd.read_excel(INPUT_FILE_NAME)
        print(f"Successfully read '{INPUT_FILE_NAME}'")
    except FileNotFoundError:
        print(f"Error: File '{INPUT_FILE_NAME}' not found!")
        return
    except Exception as e:
        print(f"Error reading Excel: {e}")
        return
    
    id_col = "Student's ID"
    name_col = "Student's Name"

    attendance_cols = df_input.columns[3:]
    #print(f"Detected attendance columns: {list(attendance_cols)}")

    results = []
    count_70_plus = count_60_to_69 = count_45_to_59 =  count_30_to_45 = count_below_30 = 0

    for index, row in df_input.iterrows():
        student_id = row[id_col]
        name = row[name_col]
        present_days = 0
        for col in attendance_cols:
            val = str(row[col]).strip().upper()
            if val == 'H' or val == 'P' or val == '1':  
                present_days += 1

        percentage = (present_days / TOTAL_CLASSES) * 100

        if percentage >= 70:
            marks = 5; count_70_plus += 1
        elif percentage >= 60:
            marks = 4; count_60_to_69 += 1
        elif percentage >= 45:
            marks = 3; count_45_to_59 += 1
        elif percentage >= 30:
            marks = 2; count_30_to_45 += 1
        else:
            marks = 0; count_below_30 += 1

        results.append({
            "Student's ID": student_id,
            "Student's Name": name,
            "Present Days": present_days,
            "Percentage (%)": round(percentage, 2),
            "Marks": marks
        })
    students_df = pd.DataFrame(results)
    students_df.index = students_df.index + 1
    students_df.index.name = 'No.'

    summary_df = pd.DataFrame({
        'Percentage Category': ['>= 70%', '>= 60% (but < 70%)', '>= 45% (but < 60%)', '>= 30% (but < 45%)', '< 30%'],
        'Student Count': [count_70_plus, count_60_to_69, count_45_to_59, count_30_to_45, count_below_30]
    })
    summary_df.index = summary_df.index + 1
    summary_df.index.name = 'No.'

    try:
        with pd.ExcelWriter(OUTPUT_FILE_NAME, engine='openpyxl') as writer:
            students_df.to_excel(writer, sheet_name='Student Marks Report')
            summary_df.to_excel(writer, sheet_name='Attendance Summary')

        print(f"\n Report generated successfully!")
        print(f" Saved as '{OUTPUT_FILE_NAME}'")
        #print(f" Columns visible: ['Student's ID', 'Student's Name', 'Present Days', 'Percentage (%)', 'Marks']")

    except Exception as e:
        print(f" Error saving file: {e}")
        print("Make sure the Excel file isn't open.")

if __name__ == "__main__":
    process_attendance_from_excel()
