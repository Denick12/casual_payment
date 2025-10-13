from datetime import datetime, timedelta

from pymysql import IntegrityError

from app import app
import pymysql
from flask_wtf.csrf import CSRFProtect
import pandas as pd
from flask import request, flash, redirect, url_for, render_template
import openpyxl
import calendar
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
                file_data = pd.read_excel(file, sheet_name=None, engine='openpyxl', skiprows=5)
                if list_category == 'Attendance':
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
                            for index, row in data.iloc[:-1].iterrows():
                                user_id = row['Staff No.']
                                payment_rate = row['Payment Rate']
                                # Now inner loop for extra columns
                                for col_index in days_data_columns:
                                    date = data.columns[col_index]  # header at that index
                                    cell_value = row.iloc[col_index]  # row value at that index
                                    status = 'X' if pd.isna(cell_value) else cell_value  # replace empty cells with 'X'

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
                    date = request.form['date']
                    date_obj = datetime.strptime(date, "%Y-%m-%d")
                    # ISO calendar returns a tuple: (year, week_number, weekday)
                    iso_year, iso_week, iso_weekday = date_obj.isocalendar()

                    # Required columns
                    required_columns = {'Staff No.', 'Tips & Incentives', 'Gross', 'Shif', 'Nssf', 'Housing Levy',
                                        'Advances', 'Overpayment', 'Pending Bills', 'Total Dedution',
                                        'Housing Levy Refund', 'Pending Bills',
                                        }
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
                                tips = row['Tips & Incentives']
                                gross = row['Gross']
                                shif = row['Shif']
                                nssf = row['Nssf']
                                housing_levy = row['Housing Levy']
                                advances = row['Advances']
                                overpayment = row['Overpayment']
                                pending_bills = row['Pending Bills']
                                total_deduction = row['Total Dedution']
                                housing_levy_refund = row['Housing Levy Refund']

                                # Insert into attendance table
                                cursor.execute("insert into payroll_summary (staff_no, tips, gross, shif, nssf, "
                                               "housing_levy, advances, overpayment, pending_bills, total_deduction, "
                                               "housing_levy_refund, week, year)"
                                               "values(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
                                               (user_id, tips, gross, shif, nssf, housing_levy, advances,
                                                overpayment, pending_bills, total_deduction, housing_levy_refund,
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


@app.route('/payroll_summary', methods=['GET', 'POST'])
def payroll_summary():
    conn, cursor = db_connection()
    # hard coded, to be made dynamic later
    year = 2025
    week = 26
    # ISO week: Sunday as first day
    first_day_of_the_week = datetime.fromisocalendar(year, week, 7)
    last_day_of_the_week = first_day_of_the_week + timedelta(days=6)

    # check whether the week crosses 2 different months
    crosses_month = first_day_of_the_week.month != last_day_of_the_week.month
    if crosses_month:
        # Get the month of that first day
        month = first_day_of_the_week.month

        # Get the last day of that month (date)
        last_day_of_month = calendar.monthrange(year, month)[1]  # returns (weekday_of_first_day, num_days_in_month)

        # Construct the date of the last day
        last_day_date = datetime(year, month, last_day_of_month)
        first_day_of_following_month = datetime(year, month + 1, 1)
        # compute the first range...
        cursor.execute("select date_format(%s, '%%m/%%e/%%Y'), user_id, sum(payment_rate) "
                       "from payroll_summary "
                       "join attendance_records on staff_no = user_id "
                       "where status = %s and date between %s and %s group by user_id ",
                       (last_day_date, 'P', first_day_of_the_week, last_day_date))
        first_range = cursor.fetchall()
        # required date format = 'date month_name 2025'
        sum_of_first_range = [last_day_date.strftime('%e %B %Y'), sum(row[1] for row in first_range)]
        # compute the second range...
        cursor.execute("select date_format(%s, '%%m/%%e/%%Y'), user_id, sum(payment_rate) from payroll_summary "
                       "join attendance_records on staff_no = user_id "
                       "where status = %s and date between %s and %s group by user_id ",
                       (last_day_of_the_week, 'P', first_day_of_following_month, last_day_of_the_week))
        second_range = cursor.fetchall()
        # required date format = 'date month_name 2025'
        sum_of_second_range = [last_day_of_the_week.strftime('%e %B %Y'), sum(row[1] for row in second_range)]
        data = {
            'crosses_month': crosses_month,
            'first_range': first_range,
            'sum_of_first_range': sum_of_first_range,
            'second_range': second_range,
            'sum_of_second_range': sum_of_second_range
        }
    return render_template('payroll_summary.html', data=data)
