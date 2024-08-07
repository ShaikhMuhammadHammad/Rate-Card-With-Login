import os
import pandas as pd
from flask import Flask, render_template, request, session, redirect, url_for, flash, send_file, make_response
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = 'bcpkey'
app.config['UPLOAD_FOLDER'] = 'uploads'  

ALLOWED_EXTENSIONS = {'xls', 'xlsx', 'csv'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

class User:
    def __init__(self, id, username, password):
        self.id = id
        self.username = username
        self.password = password

    def __repr__(self):
        return f'<User:{self.username}>'

users = []
users.append(User(id=1, username='Hammad', password='Techno@123'))

@app.route("/", methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        session.pop('user_id', None)
        username = request.form['username']
        password = request.form['password']
        user = next((x for x in users if x.username == username), None)
        if user and user.password == password:
            session['user_id'] = user.id
            return redirect(url_for('upload_files'))
        flash('Invalid username or password')
        return redirect(url_for('login'))
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.pop('user_id', None)
    flash("You have been logged out")
    return redirect(url_for('login'))

@app.route('/upload', methods=['GET', 'POST'])
def upload_files():
    if 'user_id' not in session:
        flash("Please log in to access this page")
        return redirect(url_for('login'))
    if request.method == 'POST':
        try:
            if 'file1' not in request.files or 'file2' not in request.files:
                flash("Missing file(s) in request.")
                return redirect(request.url)

            file1 = request.files['file1']
            file2 = request.files['file2']

            if file1.filename == '' or file2.filename == '':
                flash("One or both files have no filename.")
                return redirect(request.url)

            if file1 and allowed_file(file1.filename) and file2 and allowed_file(file2.filename):
                filename1 = secure_filename(file1.filename)
                filename2 = secure_filename(file2.filename)

                if not os.path.exists(app.config['UPLOAD_FOLDER']):
                    os.makedirs(app.config['UPLOAD_FOLDER'])

                file1_path = os.path.join(app.config['UPLOAD_FOLDER'], filename1)
                file2_path = os.path.join(app.config['UPLOAD_FOLDER'], filename2)

                file1.save(file1_path)
                file2.save(file2_path)

                df1 = pd.read_excel(file1_path)
                df2 = pd.read_excel(file2_path)

                try:
                    df3 = pd.read_excel(file1_path, sheet_name='Sheet1')
                    df4 = pd.read_excel(file1_path, sheet_name='Sheet2')
                except Exception as e:
                    print(f"Error loading df3 and df4: {e}")
                    df3 = pd.DataFrame()
                    df4 = pd.DataFrame()

                try:
                    df5 = pd.read_excel(file1_path, sheet_name='Sheet3')  # Adjust as necessary
                    df5.columns = df5.columns.str.strip()  # Remove trailing spaces from column names
                except Exception as e:
                    print(f"Error loading df5: {e}")
                    df5 = pd.DataFrame()

                for index, row in df2.iterrows():
                    condition = (
                        (df1['Channel'].str.lower() == row['Channel'].lower()) &
                        (row['AdStart'] >= df1['Starttime']) &
                        (row['AdEnd'] <= df1['EndTime'])
                    )
                    if condition.any():
                        matching_rate = df1.loc[condition, 'Rate'].values[0]
                        df2.at[index, 'RPM'] = matching_rate

                na_rows = pd.isnull(df2['RPM'])
                for index, row in df2[na_rows].iterrows():
                    condition = (
                        (df1['Channel'].str.lower() == row['Channel'].lower()) &
                        (row['AdStart'] <= df1['EndTime'])
                    )
                    if condition.any():
                        matching_rate = df1.loc[condition, 'Rate'].values[0]
                        df2.at[index, 'RPM'] = matching_rate

                if not df3.empty:
                    for index, row in df2.iterrows():
                        condition = (
                            (df3['Channel'].str.lower() == row['Channel'].lower()) &
                            (row['AdStart'] >= df3['StartTime']) &
                            (row['AdEnd'] <= df3['EndTime']) &
                            (df3['programName'] == row['programName']) &
                            (df3['Day'] == row['Day'])
                        )
                        if condition.any():
                            matching_rate = df3.loc[condition, 'Rate'].values[0]
                            df2.at[index, 'RPM'] = matching_rate

                if not df4.empty:
                    for index, row in df2.iterrows():
                        condition = (
                            (df4['Channel'].str.lower() == row['Channel'].lower()) &
                            (row['AdStart'] >= df4['StartTime']) &
                            (row['AdEnd'] <= df4['EndTime']) &
                            (df4['Day'] == row['Day'])
                        )
                        if condition.any():
                            matching_rate = df4.loc[condition, 'Rate'].values[0]
                            df2.at[index, 'RPM'] = matching_rate

                if not df5.empty:
                    channels_in_df5 = df5['Channel'].str.lower().unique()
                    df2['Channel_lower'] = df2['Channel'].str.lower()

                    df2.loc[df2['Channel_lower'].isin(channels_in_df5), 'RPM'] = 0

                    for index, row in df2.iterrows():
                        condition = (
                            (df5['Channel'].str.lower() == row['Channel'].lower()) &
                            (df5['programName'] == row['programName']) 
                        )
                        if condition.any():
                            matching_rate = df5.loc[condition, 'Rate'].values[0]
                            df2.at[index, 'RPM'] = matching_rate

                    df2.drop(columns=['Channel_lower'], inplace=True)

                condition = (
                    (df2['Channel'] == 'ARY DIGITAL') &
                    (df2['TransmissionHour'] >= 19) &
                    (df2['TransmissionHour'] <= 22) &
                    (df2['programName'] == 'Jeeto Pakistan')
                )
                df2.loc[condition, 'RPM'] = 255500

                condition = (
                    (df2['Channel'] == 'NEO TV') &
                    (df2['TransmissionHour'] == 23) &
                    (df2['programName'] == 'G Sarkar')
                )
                df2.loc[condition, 'RPM'] = 25000

                condition = (
                    (df2['Channel'] == 'NEO TV') &
                    (df2['TransmissionHour'] == 23) &
                    (df2['programName'] == 'Zabardast')
                )
                df2.loc[condition, 'RPM'] = 25000

                original_filename = secure_filename(file2.filename)
                output_path = os.path.join(app.config['UPLOAD_FOLDER'], original_filename)
                df2.to_excel(output_path, index=False)

                return send_file(output_path, as_attachment=True, download_name=original_filename)

            flash("One or both files are not allowed types.")
            return redirect(request.url)

        except Exception as e:
            print(f"Error occurred: {e}")
            flash(f"An error occurred during processing: {e}")
            return redirect('/')
        
    response = make_response(render_template('upload.html'))
    response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, post-check=0, pre-check=0, max-age=0'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '-1'
    return response

if __name__ == "__main__":
    app.run(debug=False, host='0.0.0.0')
