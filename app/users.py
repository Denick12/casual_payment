from datetime import datetime

from pymysql import IntegrityError

from app import app
import pymysql
from flask_wtf.csrf import CSRFProtect
import pandas as pd
from flask import request, flash, redirect, url_for, render_template
import openpyxl

csrf = CSRFProtect(app)


def db_connection():
    conn = pymysql.connect(host=app.config["DB_HOST"], user=app.config["DB_USERNAME"],
                           password=app.config["DB_PASSWORD"],
                           database=app.config["DB_NAME"])
    cursor = conn.cursor()
    return conn, cursor


def allowed_file(filename):
    return ('.' in filename and
            filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS'])


@app.route('/list_upload/<list_category>', methods=['GET', 'POST'])
def list_upload(list_category):
    if request.method == 'POST':
        conn, cursor = db_connection()
        file = request.files['list_file']
        if file.filename == '':
            flash('No selected file', 'warning')
            return redirect(url_for('list_upload', list_category=list_category))
        elif file and allowed_file(file.filename):
            try:
                if list_category == 'Attendance':
                    file_data = pd.read_excel(file, sheet_name=None, engine='openpyxl')
                    # Required columns
                    required_columns = {'Staff No.', 'Payment Rate'}
                    # Iterate over each sheet
                    for sheet_name, data in file_data.items():
                        # Check for missing columns
                        uploaded_columns = set(data.columns)
                        missing_columns = required_columns - uploaded_columns

                        if not missing_columns:
                            days_data_columns = [4, 5, 6, 7, 8, 9, 10]
                            # Iterate over each row in the sheet
                            for index, row in data.iterrows():
                                user_id = row['Staff No.']
                                payment_rate = row['Payment Rate']
                                # Now inner loop for extra columns
                                for col_index in days_data_columns:
                                    date = data.columns[col_index]  # header at that index
                                    status = row.iloc[col_index]  # row value at that index

                                    # Insert into attendance table
                                    cursor.execute(
                                        "INSERT INTO attendance_records(user_id, date, payment_rate, status) "
                                        "VALUES(%s,%s,%s,%s)", (user_id, date, payment_rate, status)
                                    )
                            conn.commit()
                            flash("Attendance list uploaded successfully", 'success')
                        else:
                            missing_columns_str = ', '.join(missing_columns)
                            flash(f"The following columns are missing from {sheet_name} sheet: {missing_columns_str}. "
                                  f"Please confirm and re-upload.", 'danger')
                elif list_category == 'Payroll Summary':
                    file_data = pd.read_excel(file, sheet_name=None, engine='openpyxl', skiprows=5)
                    date = request.form['date']
                    date_obj = datetime.strptime(date, "%Y-%m-%d")
                    # ISO calendar returns a tuple: (year, week_number, weekday)
                    iso_year, iso_week, iso_weekday = date_obj.isocalendar()

                    # Required columns
                    required_columns = {'Staff No.', 'Days Worked', 'Tips & Incentives', 'Advances', 'Pending Bills',
                                        'Overpayment'}
                    # Iterate over each sheet
                    for sheet_name, data in file_data.items():
                        # Check for missing columns
                        uploaded_columns = set(data.columns)
                        missing_columns = required_columns - uploaded_columns
                        if not missing_columns:
                            # Iterate over each row in the sheet
                            for index, row in data.iloc[:-1].iterrows():
                                # print(index)
                                user_id = row['Staff No.']
                                days_worked = row['Days Worked']
                                tips = row['Tips & Incentives']
                                advances = row['Advances']
                                pending_bills = row['Pending Bills']
                                overpayment = row['Overpayment']
                                # Insert into attendance table
                                cursor.execute("insert into payroll_summary(staff_no, days_worked, tips, advances, "
                                               "pending_bills, overpayment, week, year) values(%s,%s,%s,%s,%s,%s,%s,%s)",
                                               (user_id, days_worked, tips, advances, pending_bills, overpayment,
                                                iso_week, iso_year))
                            conn.commit()
                            flash("list uploaded successfully", 'success')
                        else:
                            missing_columns_str = ', '.join(missing_columns)
                            flash(f"The following columns are missing from {sheet_name} sheet: {missing_columns_str}. "
                                  f"Please confirm and re-upload.", 'danger')
            except IntegrityError as e:
                error_code = e.args[0]
                conn.rollback()
                if error_code == 1062:  # MySQL error code for duplicate entry
                    flash(f"Duplicate entry error: ",
                          'danger')
                else:
                    flash(f"A DB error has occurred: {e}", 'danger')
            except Exception as e:
                flash(f"{e}", 'danger')
            finally:
                cursor.close()
                conn.close()
                return redirect(url_for('list_upload', list_category=list_category))
    return render_template('users.html', list_category=list_category)