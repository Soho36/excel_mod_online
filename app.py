from flask import Flask, request, render_template, send_file
import pandas as pd
import os

app = Flask(__name__)


def process_excel_file(file_path, sheet_name='מפה לביצוע'):
    # Load the Excel file
    sheet1_data = pd.read_excel(file_path, sheet_name=sheet_name)
    print("Original Data:")
    print(sheet1_data.head())

    # Identify all action columns dynamically
    action_columns = [col for col in sheet1_data.columns if col.startswith('פעולה')]

    # Filter the relevant columns including station code
    columns_of_interest = ['מקט תחנה'] + action_columns
    relevant_data = sheet1_data[columns_of_interest]
    print("Filtered Data (Relevant Columns):")
    print(relevant_data.head())

    # Remove rows where all 'פעולה' columns are NaN
    relevant_data = relevant_data.dropna(how='all', subset=action_columns)
    print("Filtered Data (After Removing NaNs):")
    print(relevant_data.head())

    # Create a list of unique actions to produce
    unique_actions = relevant_data.melt(id_vars=['מקט תחנה'], value_vars=action_columns)
    unique_actions = unique_actions.dropna().drop_duplicates().reset_index(drop=True)
    print("Unique Actions:")
    print(unique_actions.head())

    unique_actions.columns = ['Station Code', 'Action Type', 'Action']

    # Handle 505 actions: Change 'הסרה' actions in 505 to 'סטריפ ריק'
    unique_actions['Action'] = unique_actions['Action'].apply(lambda x: 'סטריפ ריק' if 'הסרה' in x else x)

    # Remove the string 'הוספה -' from the actions except for 'הוספה ברייל -'
    unique_actions['Action'] = unique_actions['Action'].apply(lambda x: x if 'הוספה ברייל -' in x else x.replace('הוספה -', ''))

    # Split Braille actions into multiple rows
    def split_braille(action):
        if 'הוספה ברייל -' in action:
            return ['ברייל ' + part.strip() for part in action.replace('הוספה ברייל -', '').split(',')]
        else:
            return [action]

    split_actions = unique_actions['Action'].apply(split_braille).explode().reset_index(drop=True)
    print("Split Actions:")
    print(split_actions.head())

    # Clean Braille actions
    def clean_braille_actions(action):
        if 'הסרה ברייל' in action or 'הסרת ברייל' in action:
            return None  # Remove the action if it involves removing Braille
        if 'הוספה ברייל -' in action:
            return action.replace('הוספה ברייל -', '').strip()  # Remove the prefix for adding Braille
        if 'ברייל' in action:
            return action.replace('ברייל ', '').strip()  # Remove the prefix 'ברייל'
        return action

    split_actions = split_actions.apply(clean_braille_actions).dropna().reset_index(drop=True)
    print("Cleaned Actions:")
    print(split_actions.head())

    # Categorize and sort actions
    def categorize_action(action):
        if action.isdigit():  # Braille actions (only numbers)
            return 0
        elif any(char.isdigit() for char in action):  # Route actions (numbers mixed with text)
            return 1
        else:  # Other actions
            return 2

    bill_of_quantities = split_actions.value_counts().reset_index()
    bill_of_quantities.columns = ['Action', 'Quantity']
    bill_of_quantities['Category'] = bill_of_quantities['Action'].apply(categorize_action)
    bill_of_quantities = bill_of_quantities.sort_values(by=['Category', 'Action']).drop(columns=['Category']).reset_index(drop=True)
    print("Bill of Quantities:")
    print(bill_of_quantities.head())

    # Sheet 2: Static, Poly, and Fixture Actions
    # Filter rows where any action column contains 'סטטי', 'פולי', או 'מתקן'
    filtered_data = relevant_data[relevant_data[action_columns].apply(lambda x: x.str.contains('סטטי|פולי|מתקן', case=False, na=False).any(), axis=1)]
    print("Filtered Data for Static, Poly, Fixture Actions:")
    print(filtered_data.head())

    # Extract relevant columns
    filtered_data_output = filtered_data[['מקט תחנה'] + action_columns]

    # Melt the DataFrame to flatten it and keep only rows with relevant actions
    flattened_data = filtered_data_output.melt(id_vars=['מקט תחנה'], value_vars=action_columns, var_name='עמודת פעולה', value_name='פעולה').dropna().sort_values(by=['מקט תחנה', 'עמודת פעולה'])
    relevant_actions = flattened_data[flattened_data['פעולה'].str.contains('סטטי|פולי|מתקן', case=False)]
    result = relevant_actions[['מקט תחנה', 'פעולה']]
    print("Relevant Actions:")
    print(result.head())

    # Initialize counters for each type of action
    static_count = result['פעולה'].str.contains('סטטי', case=False).sum()
    poly_count = result['פעולה'].str.contains('פולי', case=False).sum()
    fixture_count = result['פעולה'].str.contains('מתקן', case=False).sum()

    # Create a summary DataFrame
    summary = pd.DataFrame({
        'סוג פעולה': ['סטטי', 'פולי', 'מתקן'],
        'כמות': [static_count, poly_count, fixture_count]
    })
    print("Summary:")
    print(summary)

    # Sheet 3: Flag Actions
    # Identify rows where any action column contains 'דגל'
    flag_actions = relevant_data[relevant_data[action_columns].apply(lambda row: row.str.contains('דגל', case=False, na=False).any(), axis=1)]
    print("Flag Actions:")
    print(flag_actions.head())

    # Melt the dataframe to have one action per row
    flag_actions_melted = flag_actions.melt(id_vars=['מקט תחנה'], value_vars=action_columns)
    flag_actions_melted = flag_actions_melted.dropna().reset_index(drop=True)

    # Filter rows to keep only those actions that contain 'דגל'
    flag_actions_filtered = flag_actions_melted[flag_actions_melted['value'].str.contains('דגל', case=False)]
    flag_actions_summary = flag_actions_filtered[['מקט תחנה', 'value']]
    flag_actions_summary.columns = ['Station Code', 'Action']
    print("Flag Actions Summary:")
    print(flag_actions_summary.head())

    # Sheet 4: Station Head Actions
    # Identify rows where any action column contains 'ראש תחנה'
    station_head_actions = relevant_data[relevant_data[action_columns].apply(lambda row: row.str.contains('ראש תחנה', case=False, na=False).any(), axis=1)]
    print("Station Head Actions:")
    print(station_head_actions.head())

    # Melt the dataframe to have one action per row
    station_head_actions_melted = station_head_actions.melt(id_vars=['מקט תחנה'], value_vars=action_columns)
    station_head_actions_melted = station_head_actions_melted.dropna().reset_index(drop=True)

    # Filter rows to keep only those actions that contain 'ראש תחנה'
    station_head_actions_filtered = station_head_actions_melted[station_head_actions_melted['value'].str.contains('ראש תחנה', case=False)]
    station_head_actions_summary = station_head_actions_filtered[['מקט תחנה', 'value']]
    station_head_actions_summary.columns = ['Station Code', 'Action']
    print("Station Head Actions Summary:")
    print(station_head_actions_summary.head())

    # Save all sheets to a new Excel file
    output_file_path = file_path.replace('.xlsx', ' כתב כמויות מלא.xlsx')
    with pd.ExcelWriter(output_file_path) as writer:
        bill_of_quantities.to_excel(writer, sheet_name='Final Sorted Bill of Quantities', index=False)
        result.to_excel(writer, sheet_name='Static Poly Fixture Actions', index=False)
        summary.to_excel(writer, sheet_name='Static Poly Fixture Summary', index=False)
        flag_actions_summary.to_excel(writer, sheet_name='Flag Actions', index=False)
        station_head_actions_summary.to_excel(writer, sheet_name='Station Head Actions', index=False)

    return output_file_path

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    files = request.files.getlist('files[]')
    sheet_name = request.form.get('sheet_name', 'מפה לביצוע')
    output_files = []

    for file in files:
        filename = file.filename
        file_path = os.path.join('uploads', filename)

        # Ensure unique filenames
        if os.path.exists(file_path):
            base, ext = os.path.splitext(filename)
            counter = 1
            while os.path.exists(file_path):
                new_filename = f"{base}_{counter}{ext}"
                file_path = os.path.join('uploads', new_filename)
                counter += 1

        file.save(file_path)
        output_file = process_excel_file(file_path, sheet_name)
        output_files.append(output_file)

    return render_template('result.html', output_files=output_files)

@app.route('/download/<filename>')
def download_file(filename):
    return send_file(os.path.join('uploads', filename), as_attachment=True)

# Add custom filter for Jinja2 to get the basename of a file
@app.template_filter('basename')
def basename_filter(s):
    return os.path.basename(s)


if __name__ == '__main__':
    if not os.path.exists('uploads'):
        os.makedirs('uploads')
    from os import environ
app.run(host='0.0.0.0', port=environ.get('PORT', 5000))

