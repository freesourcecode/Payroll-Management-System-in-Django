from django.shortcuts import render, redirect, get_object_or_404
from .models import Payroll, Base_payroll, Payroll_for_phil_asia, Gov_benefits
from hrms.models import Company_rates
from logs.models import Logs
from employee.models import Employee, Employee_hiring_details, Employee_medical, Employee_medical_contrib, Employee_company_loan, Employee_vale, Employee_canteen, Employee_canteen_contrib, Employee_sss_loan, Employee_pagibig_loan, Employee_comloan_contrib, Employee_sssloan_contrib, Employee_pagibigloan_contrib, Employee_valeloan_contrib, Employee_gatepass, Employee_gatepass_contrib
from .forms import BasePayrollForm, PayrollFullAddForm, ContributionForm, PayrollFullAddFormPhilAsia
from django.http import HttpResponse, HttpResponseRedirect
from django.contrib import messages
from django.db.models import Q
from itertools import chain
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from datetime import datetime, timedelta
import arrow

import csv
import xlwt

from django.http import JsonResponse

def get_overtime_formula(request):
    data= {
        'formula' : 22
    }
    return JsonResponse(data)


# for gov benefits
@login_required
def gov_benefits(request):

    form = BasePayrollForm(request.POST)
    if form.is_valid():
        company = form.cleaned_data['company']
        start_date = form.cleaned_data['start_date']
        end_date = form.cleaned_data['end_date']
        payment_method = form.cleaned_data['payment_method']

        Logs.objects.create(
            action=f"Gov benefits report generated. (company - {company}, start_date - {start_date} end_date - {end_date}, payment_method - {payment_method})", action_by=request.user, action_date=datetime.now())

        base_payroll = Base_payroll.objects.filter(
            company=company, start_date__gte=start_date, end_date__lte=end_date, payment_method=payment_method)
        response = HttpResponse(content_type='text/ms-excel')
        file_name = f"payslip.xls"
        response['Content-Disposition'] = f'attachment; filename="{file_name}"'
        wb = xlwt.Workbook(encoding='utf-8')
        ws = wb.add_sheet('payslip')
        style = 'align: wrap on, vert center, horiz center; font: bold on;  borders: left thin, right thin, top thin, bottom thin;'
        style_basic = 'align: wrap on, vert center, horiz center;  borders: left thin, right thin, top thin, bottom thin;'

        ws.write_merge(
            0, 0, 0, 2, f"{company} {start_date} - {end_date} ", xlwt.Style.easyxf(style))
        ws.write_merge(1, 1, 1, 1, f"Last name", xlwt.Style.easyxf(style))
        ws.write_merge(1, 1, 2, 2, f"First name", xlwt.Style.easyxf(style))
        ws.write_merge(1, 1, 3, 3, f"Middle Name", xlwt.Style.easyxf(style))
        ws.write_merge(1, 1, 4, 4, f"Rate", xlwt.Style.easyxf(style))
        ws.write_merge(1, 1, 5, 5, f"Birthday", xlwt.Style.easyxf(style))
        ws.write_merge(1, 1, 6, 6, f"SSS#", xlwt.Style.easyxf(style))
        ws.write_merge(1, 1, 7, 7, f"PHILHEALTH#", xlwt.Style.easyxf(style))
        ws.write_merge(1, 1, 8, 8, f"PAGIBIG#", xlwt.Style.easyxf(style))
        ws.write_merge(1, 1, 9, 9, f"MOBILE NO.#", xlwt.Style.easyxf(style))

        row = 3
        col = 10
        listahan = []
        emp_counter = 0
        total_days = 0
        total_deductions = 0
        base_col = 9
        payroll_dates = []
        for idx, base in enumerate(base_payroll):
            # print(f"{col} - {col+1}")
            payroll_dates.append(f"{base.start_date} - {base.end_date}")
            ws.write_merge(
                1, 1, col, col+1, f"{base.start_date} - {base.end_date}", xlwt.Style.easyxf(style))

            ws.write_merge(2, 2, col, col, f"DAYS", xlwt.Style.easyxf(style))
            ws.write_merge(2, 2, col+1, col+1, f"SSS/PH/PAG",
                           xlwt.Style.easyxf(style))
            # ws.write_merge(1, 1, col, col+1, f"{base.end_date}", xlwt.Style.easyxf(style))

            if str(company) == "phil asia":
                payroll_listhan = Payroll_for_phil_asia.objects.filter(
                    base_payroll=base.pk).order_by('employee__last_name')
            else:
                payroll_listhan = Payroll.objects.filter(base_payroll=base.pk)

            for pay in (payroll_listhan):
                # check if nagexist na sa list yung employee
                if pay.employee.pk not in listahan:

                    listahan.append(pay.employee.pk)

                    ws.write_merge(row, row, 0, 0, emp_counter +
                                   1, xlwt.Style.easyxf(style))
                    ws.write_merge(
                        row, row, 1, 1, f"{pay.employee.last_name}", xlwt.Style.easyxf(style_basic))
                    ws.write_merge(
                        row, row, 2, 2, f"{pay.employee.first_name}", xlwt.Style.easyxf(style_basic))
                    ws.write_merge(
                        row, row, 3, 3, f"{pay.employee.middle_name}", xlwt.Style.easyxf(style_basic))
                    ws.write_merge(
                        row, row, 4, 4, f"{pay.employee.get_hiring_details().rate}", xlwt.Style.easyxf(style_basic))
                    ws.write_merge(
                        row, row, 5, 5, f"{pay.employee.date_of_birth}", xlwt.Style.easyxf(style_basic))
                    ws.write_merge(
                        row, row, 6, 6, f"{pay.employee.sss_no}", xlwt.Style.easyxf(style_basic))
                    ws.write_merge(
                        row, row, 7, 7, f"{pay.employee.philhealth_no}", xlwt.Style.easyxf(style_basic))
                    ws.write_merge(
                        row, row, 8, 8, f"{pay.employee.pagibig_no}", xlwt.Style.easyxf(style_basic))
                    ws.write_merge(
                        row, row, 9, 9, f"{pay.employee.phone}", xlwt.Style.easyxf(style_basic))
                    # ws.write_merge(
                    #     row, row, 10, 10, f"{pay.base_payroll.start_date} - {pay.base_payroll.end_date} {pay.employee} - {pay.days}", xlwt.Style.easyxf(style_basic))
                    # ws.write_merge(row, row, 11, 11, pay.sss + pay.philhealth + pay.pagibig, xlwt.Style.easyxf(style_basic))
                    row += 1

                    emp_counter += 1

            col += 2

        # values
        row = 3
        for dates in payroll_dates:
            for base in base_payroll:
                if dates == f"{base.start_date} - {base.end_date}":
                    # print(f"{dates} ==  {base.start_date} - {base.end_date}")
                    if str(company) == "phil asia":
                        payroll_listhan = Payroll_for_phil_asia.objects.filter(
                            base_payroll=base.pk).order_by('employee__last_name')
                    else:
                        payroll_listhan = Payroll.objects.filter(
                            base_payroll=base.pk)

                    for tao in listahan:
                        for pays in payroll_listhan:
                            if tao == pays.employee_id:
                                print(f"{dates} - {pays.employee}")
                                deductions = pays.sss + pays.philhealth + pays.pagibig

                                if str(company) == "phil asia":
                                    Gov_benefits.objects.create(
                                        employee_id=tao, days=pays.days, deductions=deductions, start_date=base.start_date, end_date=base.end_date)
                                    ws.write_merge(
                                        row, row, base_col+1, base_col+1, f"{pays.days}", xlwt.Style.easyxf(style))
                                else:
                                    Gov_benefits.objects.create(
                                        employee_id=tao, days=pays.regular_days, deductions=deductions, start_date=base.start_date, end_date=base.end_date)
                                    ws.write_merge(
                                        row, row, base_col+1, base_col+1, f"{pays.regular_days}", xlwt.Style.easyxf(style))
                                ws.write_merge(
                                    row, row, base_col+2, base_col+2, f"{pays.sss + pays.philhealth + pays.pagibig}", xlwt.Style.easyxf(style))

                        row += 1
                    base_col += 2
                    row = 3

            # base_col = 9

        # get the gov benefits from tables
        row = 3
        for tao in listahan:
            total_days = 0
            total_deductions = 0
            for base in base_payroll:
                tao_data = Gov_benefits.objects.filter(
                    employee_id=tao, start_date=base.start_date, end_date=base.end_date).first()
                # return HttpResponse(tao_data.days)

                if tao_data:
                    total_days = total_days + tao_data.days
                    total_deductions += tao_data.deductions
            ws.write_merge(row, row, base_col+1, base_col+1,
                           f"{total_days}", xlwt.Style.easyxf(style))
            ws.write_merge(row, row, base_col+2, base_col+2,
                           f"{total_deductions}", xlwt.Style.easyxf(style))
            row += 1
            total_days = 0
            total_deductions = 0

        ws.write_merge(1, 1, col, col, f"Total", xlwt.Style.easyxf(style))
        ws.write_merge(2, 2, col, col, f"Days", xlwt.Style.easyxf(style))
        ws.write_merge(1, 1, col+1, col+1, f"Total", xlwt.Style.easyxf(style))
        ws.write_merge(2, 2, col+1, col+1, f"DEDUCTION",
                       xlwt.Style.easyxf(style))

        # delete table data gov benefits
        Gov_benefits.objects.all().delete()

        wb.save(response)
        return response


@login_required
def payroll_summary_phil_asia(request, pk):
    return HttpResponse(f"payslip phil asia {pk}")


@login_required
def payroll_payslip_phil_asia(request, pk):
    response = HttpResponse(content_type='text/ms-excel')
    file_name = f"payslip.xls"
    response['Content-Disposition'] = f'attachment; filename="{file_name}"'
    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('payslip')
    got_payroll = Payroll_for_phil_asia.objects.filter(base_payroll=pk)
    style = 'align: wrap on, vert center, horiz center; font: bold on;  borders: left thin, right thin, top thin, bottom thin;'
    style_basic = 'align: wrap on, vert center, horiz center;  borders: left thin, right thin, top thin, bottom thin;'
    row = 0
    col = 0
    emp_count = 0
    emp_count_per_row = 0
    for pay in got_payroll:
        emp_count_per_row = emp_count_per_row + 1
        emp_count = emp_count + 1
        employee_name = f"{pay.employee.last_name} {pay.employee.first_name}"
        print(employee_name)
        get_company_features = get_object_or_404(
            Company_rates, company=pay.base_payroll.company)
        row_plus = 3
        if emp_count_per_row <= 3:
            ws.write(row, col, "Name", xlwt.Style.easyxf(style))
            ws.write(row, col+1, emp_count, xlwt.Style.easyxf(style))
            # ws.write(0, 1, "1", xlwt.Style.easyxf(style))
            ws.write_merge(row, row, col+2, col+3,
                           f'{employee_name}', xlwt.Style.easyxf(style))
            ws.write_merge(row+1, row+1, col, col+3,
                           f'Pay Period {pay.base_payroll.start_date} - {pay.base_payroll.end_date}', xlwt.Style.easyxf(style))
            ws.write(row+2, col+3, "Amount", xlwt.Style.easyxf(style))

            # days
            ws.write(row+row_plus, col, "days", xlwt.Style.easyxf(style))
            ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style_basic))
            ws.write(row+row_plus, col+2, pay.days,
                     xlwt.Style.easyxf(style_basic))
            ws.write(row+row_plus, col+3, pay.amount,
                     xlwt.Style.easyxf(style_basic))
            row_plus += 1

            ws.write(row+row_plus, col, "HRS", xlwt.Style.easyxf(style))
            ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style_basic))
            ws.write(row+row_plus, col+2, pay.days,
                     xlwt.Style.easyxf(style_basic))
            ws.write(row+row_plus, col+3, pay.pay,
                     xlwt.Style.easyxf(style_basic))
            row_plus += 1

            ws.write(row+row_plus, col, "ND", xlwt.Style.easyxf(style))
            ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style_basic))
            ws.write(row+row_plus, col+2, pay.nd,
                     xlwt.Style.easyxf(style_basic))
            ws.write(row+row_plus, col+3, pay.nd_pay,
                     xlwt.Style.easyxf(style_basic))
            row_plus += 1

            ws.write(row+row_plus, col, "HRS OT", xlwt.Style.easyxf(style))
            ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style_basic))
            ws.write(row+row_plus, col+2, pay.hrs_ot,
                     xlwt.Style.easyxf(style_basic))
            ws.write(row+row_plus, col+3, pay.hrs_ot_pay,
                     xlwt.Style.easyxf(style_basic))
            row_plus += 1

            # Total amount
            ws.write(row+row_plus, col, "TOTAL AMOUNT",
                     xlwt.Style.easyxf(style))
            ws.write(row+row_plus, col+3, pay.amount + pay.pay +
                     pay.hrs_ot_pay + pay.nd_pay, xlwt.Style.easyxf(style))

            row_plus = row_plus + 2
            ws.write(row+row_plus, col, "LESS", xlwt.Style.easyxf(style))
            row_plus = row_plus + 1

            # sss:
            ws.write(row+row_plus, col, "CANTEEN", xlwt.Style.easyxf(style))
            ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style_basic))
            # ws.write(row+row_plus, col+2, pay.uniform, xlwt.Style.easyxf(style_basic))
            ws.write(row+row_plus, col+3, pay.sss,
                     xlwt.Style.easyxf(style_basic))
            row_plus += 1
 # sss:
            ws.write(row+row_plus, col, "OFFICE", xlwt.Style.easyxf(style))
            ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style_basic))
            # ws.write(row+row_plus, col+2, pay.uniform, xlwt.Style.easyxf(style_basic))
            ws.write(row+row_plus, col+3, pay.sss,
                     xlwt.Style.easyxf(style_basic))
            row_plus += 1
 # sss:
            ws.write(row+row_plus, col, "SSS", xlwt.Style.easyxf(style))
            ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style_basic))
            # ws.write(row+row_plus, col+2, pay.uniform, xlwt.Style.easyxf(style_basic))
            ws.write(row+row_plus, col+3, pay.sss,
                     xlwt.Style.easyxf(style_basic))
            row_plus += 1

            # philhealth
            ws.write(row+row_plus, col, "PHIL.", xlwt.Style.easyxf(style))
            ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style_basic))
            # ws.write(row+row_plus, col+2, pay.uniform, xlwt.Style.easyxf(style_basic))
            ws.write(row+row_plus, col+3, pay.philhealth,
                     xlwt.Style.easyxf(style_basic))
            row_plus += 1

            # pagibig
            ws.write(row+row_plus, col, "PAGIBIG", xlwt.Style.easyxf(style))
            ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style_basic))
            # ws.write(row+row_plus, col+2, pay.uniform, xlwt.Style.easyxf(style_basic))
            ws.write(row+row_plus, col+3, pay.pagibig,
                     xlwt.Style.easyxf(style_basic))
            row_plus += 1
            ws.write_merge(row+row_plus, row+row_plus, col, col+1,
                           f'NET AMOUNT', xlwt.Style.easyxf(style))
            ws.write(row+row_plus, col+3, pay.net_amount,
                     xlwt.Style.easyxf(style_basic))
            print(f"net amount: {pay.net_amount}")
        # row = row + 1
            col = col + 4

            if emp_count_per_row == 3:
                emp_count_per_row = 0
                col = 0
                row += 18

    wb.save(response)
    return response


@login_required
def payroll_payslip(requset, pk):
    response = HttpResponse(content_type='text/ms-excel')
    file_name = f"payslip.xls"
    response['Content-Disposition'] = f'attachment; filename="{file_name}"'
    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('payslip')
    got_payroll = Payroll.objects.filter(base_payroll=pk)
    style = 'align: wrap on, vert center, horiz center; font: bold on;  borders: left thin, right thin, top thin, bottom thin;'
    row = 0
    col = 0
    emp_count = 0
    emp_count_per_row = 0
    for pay in got_payroll:
        emp_count_per_row = emp_count_per_row + 1
        emp_count = emp_count + 1
        employee_name = f"{pay.employee.last_name} {pay.employee.first_name}"
        print(employee_name)
        get_company_features = get_object_or_404(
            Company_rates, company=pay.base_payroll.company)
        row_plus = 3
        if emp_count_per_row <= 3:
            ws.write(row, col, "Name", xlwt.Style.easyxf(style))
            ws.write(row, col+1, emp_count, xlwt.Style.easyxf(style))
            # ws.write(0, 1, "1", xlwt.Style.easyxf(style))
            ws.write_merge(row, row, col+2, col+3,
                           f'{employee_name}', xlwt.Style.easyxf(style))
            ws.write_merge(row+1, row+1, col, col+3,
                           f'Payroll Cov. {pay.base_payroll.start_date} - {pay.base_payroll.end_date}', xlwt.Style.easyxf(style))
            ws.write(row+2, col+3, "Amount", xlwt.Style.easyxf(style))

            # days
            ws.write(row+row_plus, col, "days", xlwt.Style.easyxf(style))
            ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style))
            ws.write(row+row_plus, col+2, pay.regular_days,
                     xlwt.Style.easyxf(style))
            ws.write(row+row_plus, col+3, pay.regular_amount,
                     xlwt.Style.easyxf(style))
            row_plus += 1

            # ecola
            if get_company_features.activate_ecola:
                ws.write(row+row_plus, col, "ECOLA", xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style))
                # # ws.write(row+row_plus, col+2, pay.regular_days, xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+3, pay.ecola,
                         xlwt.Style.easyxf(style))
                row_plus += 1

            # OT
            if get_company_features.activate_overtime:
                ws.write(row+row_plus, col, "OT", xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+2, pay.overtime_regular,
                         xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+3, pay.overtime_regular_amount,
                         xlwt.Style.easyxf(style))
                row_plus += 1

            # SUNDAY
            if get_company_features.activate_sunday:
                ws.write(row+row_plus, col, "SUNDAY", xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+2, pay.sunday,
                         xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+3, pay.sunday_amount,
                         xlwt.Style.easyxf(style))
                row_plus += 1

            # SUNDAY OT
                ws.write(row+row_plus, col, "SUN. OT",
                         xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+2, pay.sunday_overtime,
                         xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+3, pay.sunday_overtime_amount,
                         xlwt.Style.easyxf(style))
                row_plus += 1
            # SUNDAY ND
                ws.write(row+row_plus, col, "SUN. ND",
                         xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+2, pay.sunday_nd,
                         xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+3, pay.sunday_nd_amount,
                         xlwt.Style.easyxf(style))
                row_plus += 1
            # Holiday

            if get_company_features.activate_holiday:
                ws.write(row+row_plus, col, "HOLIDAY",
                         xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+2, pay.holiday_regular_days,
                         xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+3, pay.holiday_regular_amount,
                         xlwt.Style.easyxf(style))
                row_plus += 1
            # Holiday OT
                ws.write(row+row_plus, col, "HOL. OT",
                         xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+2, pay.holiday_overtime,
                         xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+3, pay.holiday_overtime_amount,
                         xlwt.Style.easyxf(style))
                row_plus += 1

            # Special Holiday
            if get_company_features.activate_special:
                ws.write(row+row_plus, col, "SP. HOL.",
                         xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+2, pay.special_holiday_days,
                         xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+3, pay.special_holiday_amount,
                         xlwt.Style.easyxf(style))
                row_plus += 1

            # Special Holiday OT
                ws.write(row+row_plus, col, "SP. HOL. OT",
                         xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+2, pay.special_holiday_overtime,
                         xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+3, pay.special_holiday_overtime_amount,
                         xlwt.Style.easyxf(style))
                row_plus += 1

            # Restday
            if get_company_features.activate_rest_day:
                ws.write(row+row_plus, col, "RD", xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+2, pay.rest_days,
                         xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+3, pay.rest_amount,
                         xlwt.Style.easyxf(style))
                row_plus += 1

                # Restday OT
                ws.write(row+row_plus, col, "RD OT", xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+2, pay.rest_day_overtime,
                         xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+3, pay.rest_day_overtime_amount,
                         xlwt.Style.easyxf(style))
                row_plus += 1

            # ND
            if get_company_features.activate_night_differential:
                ws.write(row+row_plus, col, "N. DIFF",
                         xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+2, pay.night_diff_days,
                         xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+3, pay.night_diff_amount,
                         xlwt.Style.easyxf(style))
                row_plus += 1
            print(row_plus)
            # Total amount
            ws.write(row+row_plus, col, "TOTAL AMOUNT",
                     xlwt.Style.easyxf(style))
            ws.write(row+row_plus, col+3, pay.regular_amount + pay.ecola + pay.overtime_regular_amount + pay.sunday_amount + pay.sunday_overtime_amount + pay.sunday_nd_amount + pay.holiday_regular_amount +
                     pay.holiday_overtime_amount + pay.special_holiday_amount + pay.special_holiday_overtime_amount + pay.rest_amount + pay.rest_day_overtime_amount + pay.night_diff_amount, xlwt.Style.easyxf(style))

            row_plus = row_plus + 2
            ws.write(row+row_plus, col, "LESS", xlwt.Style.easyxf(style))
            row_plus = row_plus + 1

            #tardiness / late
            if get_company_features.activate_tardiness:
                ws.write(row+row_plus, col, "Late / Und.",
                         xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+2, pay.tardiness_undertime_regular,
                         xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+3, pay.tardiness_undertime_regular_amount,
                         xlwt.Style.easyxf(style))
                row_plus += 1
            # uniform
            if get_company_features.activate_uniform:
                ws.write(row+row_plus, col, "UNI.", xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style))
                # ws.write(row+row_plus, col+2, pay.uniform, xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+3, pay.uniform,
                         xlwt.Style.easyxf(style))
                row_plus += 1

            # MEDICAL
            if get_company_features.activate_medical:
                ws.write(row+row_plus, col, "MEDICAL",
                         xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style))
                # ws.write(row+row_plus, col+2, pay.uniform, xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+3, pay.medical,
                         xlwt.Style.easyxf(style))
                row_plus += 1

            # CANTEEN
            if get_company_features.activate_canteen:
                ws.write(row+row_plus, col, "CANTEEN",
                         xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style))
                # ws.write(row+row_plus, col+2, pay.uniform, xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+3, pay.canteen,
                         xlwt.Style.easyxf(style))
                row_plus += 1

            # GATEPASS
            if get_company_features.activate_gatepass:
                ws.write(row+row_plus, col, "GATEPASS",
                         xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style))
                # ws.write(row+row_plus, col+2, pay.uniform, xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+3, pay.gatepass,
                         xlwt.Style.easyxf(style))
                row_plus += 1

            # VALE
            if get_company_features.activate_vale:
                ws.write(row+row_plus, col, "VALE", xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style))
                # ws.write(row+row_plus, col+2, pay.uniform, xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+3, pay.vale,
                         xlwt.Style.easyxf(style))
                row_plus += 1

            # 13 MONTH
            if get_company_features.activate_thirteenth_month:
                ws.write(row+row_plus, col, "13TH MON.",
                         xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style))
                # ws.write(row+row_plus, col+2, pay.uniform, xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+3, pay.thirteenth_month,
                         xlwt.Style.easyxf(style))
                row_plus += 1

            # SIL
            if get_company_features.activate_sil:
                ws.write(row+row_plus, col, "SIL", xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style))
                # ws.write(row+row_plus, col+2, pay.uniform, xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+3, pay.sil,
                         xlwt.Style.easyxf(style))
                row_plus += 1

            # TSHIRT
            if get_company_features.activate_tshirt:
                ws.write(row+row_plus, col, "TSHIRT", xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style))
                # ws.write(row+row_plus, col+2, pay.uniform, xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+3, pay.tshirt,
                         xlwt.Style.easyxf(style))
                row_plus += 1

            # RF
            if get_company_features.activate_rf:
                ws.write(row+row_plus, col, "RF", xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style))
                # ws.write(row+row_plus, col+2, pay.uniform, xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+3, pay.rf, xlwt.Style.easyxf(style))
                row_plus += 1

            # house
            if get_company_features.activate_house:
                ws.write(row+row_plus, col, "HOUSE", xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style))
                # ws.write(row+row_plus, col+2, pay.uniform, xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+3, pay.house,
                         xlwt.Style.easyxf(style))
                row_plus += 1

            # MISC
            if get_company_features.activate_misc:
                ws.write(row+row_plus, col, "MISC", xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style))
                # ws.write(row+row_plus, col+2, pay.uniform, xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+3, pay.misc,
                         xlwt.Style.easyxf(style))
                row_plus += 1

            # PANTS
            if get_company_features.activate_pants:
                ws.write(row+row_plus, col, "PANTS", xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style))
                # ws.write(row+row_plus, col+2, pay.uniform, xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+3, pay.pants,
                         xlwt.Style.easyxf(style))
                row_plus += 1

            # SERVICE FEE
            if get_company_features.activate_service_fee:
                ws.write(row+row_plus, col, "SERV.FEE",
                         xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style))
                # ws.write(row+row_plus, col+2, pay.uniform, xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+3, pay.service_fee,
                         xlwt.Style.easyxf(style))
                row_plus += 1

            # COMPANY LOANs
            if get_company_features.activate_company_loan:
                ws.write(row+row_plus, col, "COMLOAN",
                         xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style))
                # ws.write(row+row_plus, col+2, pay.uniform, xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+3, pay.company_loan,
                         xlwt.Style.easyxf(style))
                row_plus += 1

            # sss loan
            if get_company_features.activate_sss_loan:
                ws.write(row+row_plus, col, "SSSLOAN",
                         xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style))
                # ws.write(row+row_plus, col+2, pay.uniform, xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+3, pay.sss_loan,
                         xlwt.Style.easyxf(style))
                row_plus += 1

            # PAGIBIG LOAN
            if get_company_features.activate_pagibig_loan:
                ws.write(row+row_plus, col, "PAG.LOAN",
                         xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style))
                # ws.write(row+row_plus, col+2, pay.uniform, xlwt.Style.easyxf(style))
                ws.write(row+row_plus, col+3, pay.pagibig_loan,
                         xlwt.Style.easyxf(style))
                row_plus += 1
            # sss:
            ws.write(row+row_plus, col, "SSS", xlwt.Style.easyxf(style))
            ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style))
            # ws.write(row+row_plus, col+2, pay.uniform, xlwt.Style.easyxf(style))
            ws.write(row+row_plus, col+3, pay.sss, xlwt.Style.easyxf(style))
            row_plus += 1

            # philhealth
            ws.write(row+row_plus, col, "PHIL.", xlwt.Style.easyxf(style))
            ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style))
            # ws.write(row+row_plus, col+2, pay.uniform, xlwt.Style.easyxf(style))
            ws.write(row+row_plus, col+3, pay.philhealth,
                     xlwt.Style.easyxf(style))
            row_plus += 1

            # pagibig
            ws.write(row+row_plus, col, "PAGIBIG", xlwt.Style.easyxf(style))
            ws.write(row+row_plus, col+1, "-", xlwt.Style.easyxf(style))
            # ws.write(row+row_plus, col+2, pay.uniform, xlwt.Style.easyxf(style))
            ws.write(row+row_plus, col+3, pay.pagibig,
                     xlwt.Style.easyxf(style))
            row_plus += 1
            ws.write_merge(row+row_plus, row+row_plus, col, col+1,
                           f'NET AMOUNT', xlwt.Style.easyxf(style))
            ws.write(row+row_plus, col+3, pay.net_amount,
                     xlwt.Style.easyxf(style))

        # row = row + 1
            col = col + 4
            # if emp_count_per_row == 3:
            emp_count_per_row = 0
            #     col = 0
            #     row += 18

    wb.save(response)
    return response


def get_header_column(contribution_kind):
    # pagibig
    if contribution_kind == "pagibig":
        columns = ['Pagibig No.', 'Emp. ID', 'Last Name', ' First Name', 'Middle Name',
                   'Emp. Contri.', 'emplyrs. Contri', 'TIN', 'Date of Birth'
                   ]

    if contribution_kind == "sss":
        columns = ['SSS No.', 'Emp. ID', 'Last Name', ' First Name', 'Middle Name',
                   'Emp. Contri.', 'emplyrs. Contri', 'TIN', 'Date of Birth'
                   ]

    if contribution_kind == "philhealth":
        columns = ['Philhealth No.', 'Emp. ID', 'Last Name', ' First Name', 'Middle Name',
                   'Emp. Contri.', 'emplyrs. Contri', 'TIN', 'Date of Birth'
                   ]

    if contribution_kind == "vale":
        columns = ['Emp. ID', 'Last Name', ' First Name', 'Middle Name',
                   'Emp. Contri.', 'Date of Birth'
                   ]

    if contribution_kind == "company_loan":
        columns = ['Emp. ID', 'Last Name', ' First Name', 'Middle Name',
                   'Emp. Contri.', 'Date of Birth'
                   ]

    if contribution_kind == "sss_loan":
        columns = ['Emp. ID', 'Last Name', ' First Name', 'Middle Name',
                   'Emp. Contri.', 'Date of Birth'
                   ]

    if contribution_kind == "pagibig_loan":
        columns = ['Emp. ID', 'Last Name', ' First Name', 'Middle Name',
                   'Emp. Contri.', 'Date of Birth'
                   ]

    return columns


def get_payroll_values(company, base_payroll_id, contribution_kind):
    # return HttpResponse("eto yun")
    # lipat sa employees yung pagibig, sss, philhealth number
    if contribution_kind == "pagibig":
        # check if phil-asia
        payroll_record = Payroll.objects.filter(base_payroll=base_payroll_id).values_list(
            'employee__pagibig', 'employee__emp_id', 'employee__last_name', 'employee__first_name', 'employee__middle_name', 'pagibig', 'pagibig_employer',
            'employee__tin', 'employee__date_of_birth')
    # employer contribution sss, phil, pagibig
    # do sss
    if contribution_kind == "sss":
        return HttpResponse(company)
        if company == "phil asia":
            payroll_record = Payroll_for_phil_asia.objects.filter(base_payroll=base_payroll_id).values_list(
                'employee__sss', 'employee__emp_id', 'employee__last_name', 'employee__first_name', 'employee__middle_name', 'sss', 'sss_employer',
                'employee__tin', 'employee__date_of_birth')
        else:
            payroll_record = Payroll.objects.filter(base_payroll=base_payroll_id).values_list(
                'employee__sss', 'employee__emp_id', 'employee__last_name', 'employee__first_name', 'employee__middle_name', 'sss', 'sss_employer',
                'employee__tin', 'employee__date_of_birth')
    # do philhealth
    if contribution_kind == "philhealth":
        payroll_record = Payroll.objects.filter(base_payroll=base_payroll_id).values_list(
            'employee__philhealth', 'employee__emp_id', 'employee__last_name', 'employee__first_name', 'employee__middle_name', 'philhealth', 'philhealth_employer',
            'employee__tin', 'employee__date_of_birth')

    # do vale
    if contribution_kind == "company_loan":
        payroll_record = Payroll.objects.filter(base_payroll=base_payroll_id).values_list(
            'employee__emp_id', 'employee__last_name', 'employee__first_name', 'employee__middle_name', 'company_loan', 'employee__date_of_birth')

    # do company loan
    if contribution_kind == "vale":
        payroll_record = Payroll.objects.filter(base_payroll=base_payroll_id).values_list(
            'employee__emp_id', 'employee__last_name', 'employee__first_name', 'employee__middle_name', 'vale', 'employee__date_of_birth')
    # do sss loan
    if contribution_kind == "sss_loan":
        payroll_record = Payroll.objects.filter(base_payroll=base_payroll_id).values_list(
            'employee__emp_id', 'employee__last_name', 'employee__first_name', 'employee__middle_name', 'sss_loan', 'employee__date_of_birth')
    # do pagibig loan
    if contribution_kind == "pagibig_loan":
        payroll_record = Payroll.objects.filter(base_payroll=base_payroll_id).values_list(
            'employee__emp_id', 'employee__last_name', 'employee__first_name', 'employee__middle_name', 'pagibig_loan', 'employee__date_of_birth')
    return payroll_record


# @login_required
def payroll_contributions(request):

    if request.method == "POST":
        form = ContributionForm(request.POST)
        if form.is_valid():
            company = form.cleaned_data['company']
            start_date = form.cleaned_data['start_date']
            end_date = form.cleaned_data['end_date']
            payment_method = form.cleaned_data['payment_method']
            contribution_kind = form.cleaned_data['contribution_kind']
            # return HttpResponse(form)
            # search the base payroll for start date and end date

            # put employees in excel
            response = HttpResponse(content_type='text/ms-excel')
            file_name = f"{company} - {contribution_kind} contribution.xls"
            response['Content-Disposition'] = f'attachment; filename="{file_name}"'

            wb = xlwt.Workbook(encoding='utf-8')
            ws = wb.add_sheet(f'{contribution_kind}')

            style = 'align: wrap on, vert center, horiz center; font: bold on'
            head_title = f"Luxor Manpower Corporation - {company} {contribution_kind} contribution"
            ws.write_merge(0, 1, 0, 10, head_title, xlwt.Style.easyxf(style))

            row_num = 4

            style = 'align: wrap on, vert centre, horiz center; font: bold on'
            font_style = xlwt.Style.easyxf(style)

            columns = get_header_column(f"{contribution_kind}")

            print(columns)
            for col_num in range(len(columns)):
                ws.write(row_num, col_num, columns[col_num], font_style)

            # write the employees
            style = 'align: wrap on, vert centre, horiz center'
            # font_style = xlwt.Style.easyxf(style, num_format_str='YY-MM-DD')
            font_style = xlwt.Style.easyxf(style)
            # print (len(columns))
            col_num = 0
            emp_list = []

            base_payroll_used = Base_payroll.objects.filter(
                company=company, start_date__gte=start_date, end_date__lte=end_date, payment_method=payment_method).first()
            # return HttpResponse(f"eto yun:  {base_payroll_used.id}")
            if base_payroll_used == None:
                messages.error(request, f'Payroll not available')
                return redirect('payroll-contributions',)

            test_list = []
            # for ip in included_payrolls:
            payroll_record = get_payroll_values(
                company, base_payroll_used.id, contribution_kind)
            # total = payroll_record.count()
            # return HttpResponse(total)
            print(payroll_record)
            for row in payroll_record:
                row_num += 1
               # print(f"somethin - {row.employee}")
                for col_num in range(len(row)):
                    print(contribution_kind)
                    if contribution_kind == "sss_loan" or contribution_kind == "pagibig_loan" or contribution_kind == "company_loan" or contribution_kind == "vale":
                        if col_num == 5:
                            font_style = xlwt.Style.easyxf(
                                style, num_format_str='YY-MM-DD')
                        else:
                            font_style = xlwt.Style.easyxf(style)
                    else:
                        if col_num == 8:
                            font_style = xlwt.Style.easyxf(
                                style, num_format_str='YY-MM-DD')
                        else:
                            font_style = xlwt.Style.easyxf(style)
                    ws.write(row_num, col_num, row[col_num], font_style)
                    print(row[col_num])
            wb.save(response)
            return response

    else:
        form = ContributionForm()
    context = {
        'title': 'company',
        'head': f'Payroll Contributions',
        'form': form
    }

    return render(request, 'payrolllist/payroll_contributions.html', context)


@login_required
def payroll_billing_phil_asia(request, pk):
    response = HttpResponse(content_type='text/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="billing.xls"'

    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('payroll')

    # heading
    # get
    base_payroll = get_object_or_404(Base_payroll, pk=pk)
    # print(base_payroll)
    payrolllist = Payroll_for_phil_asia.objects.filter(
        base_payroll=base_payroll)
    style = 'align: wrap on, horiz center, vert center; font: bold on'
    style_basic = 'align: wrap on, horiz center, vert center; '
#    ws.row(0).write(0, value, xlwt.Style.easyxf(style))
    ws.write_merge(0, 0, 0, 1, 'LUXOR MANPOWER PAYROLL',
                   xlwt.Style.easyxf(style))
    ws.write_merge(1, 1, 0, 1, 'FIRM NAME: PHIL ASIA',
                   xlwt.Style.easyxf(style))
    ws.write_merge(
        2, 2, 0, 3, f'Payroll Covered: {base_payroll}', xlwt.Style.easyxf(style))

    helpers_count = Employee_hiring_details.objects.filter(
        payment_method=base_payroll.payment_method, position="helper").count()
    operators_count = Employee_hiring_details.objects.filter(
        payment_method=base_payroll.payment_method, position="operator").count()
    special_operators_count = Employee_hiring_details.objects.filter(
        payment_method=base_payroll.payment_method, position="special_operator").count()
    drivers_count = Employee_hiring_details.objects.filter(
        payment_method=base_payroll.payment_method, position="driver").count()

    ws.write_merge(
        4, 5, 0, 1, f'Helper', xlwt.Style.easyxf(style))
    ws.write_merge(
        4, 5, 2, 2, f'Days', xlwt.Style.easyxf(style))
    ws.write_merge(
        4, 5, 3, 3, f'Reg. HRS', xlwt.Style.easyxf(style))
    ws.write_merge(
        4, 5, 4, 4, f'ND', xlwt.Style.easyxf(style))
    ws.write_merge(
        4, 5, 5, 5, f'REG. OT', xlwt.Style.easyxf(style))

    rowrow = 5
    total_days = 0
    total_hrs = 0
    total_nd = 0
    total_hrs_ot = 0
    for idx, listahan in enumerate(payrolllist):

        # return HttpResponse(listahan.employee)
        if "helper" == listahan.employee.get_hiring_details().position:
            rowrow += 1
            ws.write_merge(
                rowrow, rowrow, 0, 0, idx+1, xlwt.Style.easyxf(style_basic))
            ws.write_merge(
                rowrow, rowrow, 1, 1, f"{listahan.employee}", xlwt.Style.easyxf(style_basic))
            ws.write_merge(
                rowrow, rowrow, 2, 2, f"{listahan.days}", xlwt.Style.easyxf(style_basic))
            ws.write_merge(
                rowrow, rowrow, 3, 3, f"{listahan.hrs}", xlwt.Style.easyxf(style_basic))
            ws.write_merge(
                rowrow, rowrow, 4, 4, f"{listahan.nd}", xlwt.Style.easyxf(style_basic))
            ws.write_merge(
                rowrow, rowrow, 5, 5, f"{listahan.hrs_ot}", xlwt.Style.easyxf(style_basic))
            total_days += listahan.days
            total_hrs += listahan.hrs
            total_nd += listahan.nd
            total_hrs_ot += listahan.hrs_ot

    rowrow += 2
    ws.write_merge(
        rowrow, rowrow, 0, 1, f"SUB - TOTAL", xlwt.Style.easyxf(style))
    ws.write_merge(
        rowrow, rowrow, 2, 2, f"{total_days}", xlwt.Style.easyxf(style))
    ws.write_merge(
        rowrow, rowrow, 3, 3, f"{total_hrs}", xlwt.Style.easyxf(style))
    ws.write_merge(
        rowrow, rowrow, 4, 4, f"{total_nd}", xlwt.Style.easyxf(style))
    ws.write_merge(
        rowrow, rowrow, 5, 5, f"{total_hrs_ot}", xlwt.Style.easyxf(style))

    rowrow += 2

    ws.write_merge(
        rowrow, rowrow+1, 0, 1, f'Operator', xlwt.Style.easyxf(style))
    ws.write_merge(
        rowrow, rowrow+1, 2, 2, f'Days', xlwt.Style.easyxf(style))
    ws.write_merge(
        rowrow, rowrow+1, 3, 3, f'Reg. HRS', xlwt.Style.easyxf(style))
    ws.write_merge(
        rowrow, rowrow+1, 4, 4, f'ND', xlwt.Style.easyxf(style))
    ws.write_merge(
        rowrow, rowrow+1, 5, 5, f'REG. OT', xlwt.Style.easyxf(style))

    total_days = 0
    total_hrs = 0
    total_nd = 0
    total_hrs_ot = 0

    rowrow += 1
    for idx, listahan in enumerate(payrolllist):

        # return HttpResponse(listahan.employee)
        if "operator" == listahan.employee.get_hiring_details().position:
            rowrow += 1
            ws.write_merge(
                rowrow, rowrow, 0, 0, idx+1, xlwt.Style.easyxf(style_basic))
            ws.write_merge(
                rowrow, rowrow, 1, 1, f"{listahan.employee}", xlwt.Style.easyxf(style_basic))
            ws.write_merge(
                rowrow, rowrow, 2, 2, f"{listahan.days}", xlwt.Style.easyxf(style_basic))
            ws.write_merge(
                rowrow, rowrow, 3, 3, f"{listahan.hrs}", xlwt.Style.easyxf(style_basic))
            ws.write_merge(
                rowrow, rowrow, 4, 4, f"{listahan.nd}", xlwt.Style.easyxf(style_basic))
            ws.write_merge(
                rowrow, rowrow, 5, 5, f"{listahan.hrs_ot}", xlwt.Style.easyxf(style_basic))
            total_days += listahan.days
            total_hrs += listahan.hrs
            total_nd += listahan.nd
            total_hrs_ot += listahan.hrs_ot

    rowrow += 2
    ws.write_merge(
        rowrow, rowrow, 0, 1, f"SUB - TOTAL", xlwt.Style.easyxf(style))
    ws.write_merge(
        rowrow, rowrow, 2, 2, f"{total_days}", xlwt.Style.easyxf(style))
    ws.write_merge(
        rowrow, rowrow, 3, 3, f"{total_hrs}", xlwt.Style.easyxf(style))
    ws.write_merge(
        rowrow, rowrow, 4, 4, f"{total_nd}", xlwt.Style.easyxf(style))
    ws.write_merge(
        rowrow, rowrow, 5, 5, f"{total_hrs_ot}", xlwt.Style.easyxf(style))

    rowrow += 2

    ws.write_merge(
        rowrow, rowrow+1, 0, 1, f'Special Operator', xlwt.Style.easyxf(style))
    ws.write_merge(
        rowrow, rowrow+1, 2, 2, f'Days', xlwt.Style.easyxf(style))
    ws.write_merge(
        rowrow, rowrow+1, 3, 3, f'Reg. HRS', xlwt.Style.easyxf(style))
    ws.write_merge(
        rowrow, rowrow+1, 4, 4, f'ND', xlwt.Style.easyxf(style))
    ws.write_merge(
        rowrow, rowrow+1, 5, 5, f'REG. OT', xlwt.Style.easyxf(style))

    total_days = 0
    total_hrs = 0
    total_nd = 0
    total_hrs_ot = 0

    rowrow += 1
    for idx, listahan in enumerate(payrolllist):

        # return HttpResponse(listahan.employee)
        if "special_operator" == listahan.employee.get_hiring_details().position:
            rowrow += 1
            ws.write_merge(
                rowrow, rowrow, 0, 0, idx+1, xlwt.Style.easyxf(style_basic))
            ws.write_merge(
                rowrow, rowrow, 1, 1, f"{listahan.employee}", xlwt.Style.easyxf(style_basic))
            ws.write_merge(
                rowrow, rowrow, 2, 2, f"{listahan.days}", xlwt.Style.easyxf(style_basic))
            ws.write_merge(
                rowrow, rowrow, 3, 3, f"{listahan.hrs}", xlwt.Style.easyxf(style_basic))
            ws.write_merge(
                rowrow, rowrow, 4, 4, f"{listahan.nd}", xlwt.Style.easyxf(style_basic))
            ws.write_merge(
                rowrow, rowrow, 5, 5, f"{listahan.hrs_ot}", xlwt.Style.easyxf(style_basic))
            total_days += listahan.days
            total_hrs += listahan.hrs
            total_nd += listahan.nd
            total_hrs_ot += listahan.hrs_ot

    rowrow += 2
    ws.write_merge(
        rowrow, rowrow, 0, 1, f"SUB - TOTAL", xlwt.Style.easyxf(style))
    ws.write_merge(
        rowrow, rowrow, 2, 2, f"{total_days}", xlwt.Style.easyxf(style))
    ws.write_merge(
        rowrow, rowrow, 3, 3, f"{total_hrs}", xlwt.Style.easyxf(style))
    ws.write_merge(
        rowrow, rowrow, 4, 4, f"{total_nd}", xlwt.Style.easyxf(style))
    ws.write_merge(
        rowrow, rowrow, 5, 5, f"{total_hrs_ot}", xlwt.Style.easyxf(style))

    rowrow += 2

    ws.write_merge(
        rowrow, rowrow+1, 0, 1, f'Driver', xlwt.Style.easyxf(style))
    ws.write_merge(
        rowrow, rowrow+1, 2, 2, f'Days', xlwt.Style.easyxf(style))
    ws.write_merge(
        rowrow, rowrow+1, 3, 3, f'Reg. HRS', xlwt.Style.easyxf(style))
    ws.write_merge(
        rowrow, rowrow+1, 4, 4, f'ND', xlwt.Style.easyxf(style))
    ws.write_merge(
        rowrow, rowrow+1, 5, 5, f'REG. OT', xlwt.Style.easyxf(style))

    total_days = 0
    total_hrs = 0
    total_nd = 0
    total_hrs_ot = 0

    rowrow += 1
    for idx, listahan in enumerate(payrolllist):

        # return HttpResponse(listahan.employee)
        if "driver" == listahan.employee.get_hiring_details().position:
            rowrow += 1
            ws.write_merge(
                rowrow, rowrow, 0, 0, idx+1, xlwt.Style.easyxf(style_basic))
            ws.write_merge(
                rowrow, rowrow, 1, 1, f"{listahan.employee}", xlwt.Style.easyxf(style_basic))
            ws.write_merge(
                rowrow, rowrow, 2, 2, f"{listahan.days}", xlwt.Style.easyxf(style_basic))
            ws.write_merge(
                rowrow, rowrow, 3, 3, f"{listahan.hrs}", xlwt.Style.easyxf(style_basic))
            ws.write_merge(
                rowrow, rowrow, 4, 4, f"{listahan.nd}", xlwt.Style.easyxf(style_basic))
            ws.write_merge(
                rowrow, rowrow, 5, 5, f"{listahan.hrs_ot}", xlwt.Style.easyxf(style_basic))
            total_days += listahan.days
            total_hrs += listahan.hrs
            total_nd += listahan.nd
            total_hrs_ot += listahan.hrs_ot

    rowrow += 2
    ws.write_merge(
        rowrow, rowrow, 0, 1, f"SUB - TOTAL", xlwt.Style.easyxf(style))
    ws.write_merge(
        rowrow, rowrow, 2, 2, f"{total_days}", xlwt.Style.easyxf(style))
    ws.write_merge(
        rowrow, rowrow, 3, 3, f"{total_hrs}", xlwt.Style.easyxf(style))
    ws.write_merge(
        rowrow, rowrow, 4, 4, f"{total_nd}", xlwt.Style.easyxf(style))
    ws.write_merge(
        rowrow, rowrow, 5, 5, f"{total_hrs_ot}", xlwt.Style.easyxf(style))

    wb.save(response)
    return response


@login_required
def payroll_billing(request, pk):
    response = HttpResponse(content_type='text/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="billing.xls"'

    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('billing')

    # heading
    # get
    base_payroll = get_object_or_404(Base_payroll, pk=pk)
    print(base_payroll)
    style = 'align: wrap on, vert centre, horiz center; font: bold on'
#    ws.row(0).write(0, value, xlwt.Style.easyxf(style))
    ws.write_merge(0, 1, 0, 20, 'Luxor Manpower Payroll',
                   xlwt.Style.easyxf(style))
    ws.write_merge(
        2, 2, 0, 4, f'Payroll Covered: {base_payroll}', xlwt.Style.easyxf(style))
    # ws.write_merge(1, 0, 0, 0, 'Payroll Covered', xlwt.Style.easyxf(style))
    # Sheet header, first row

    row_num = 4

    style = 'align: wrap on, vert centre, horiz center; font: bold on'
    font_style = xlwt.Style.easyxf(style)
    columns = ['Employee', 'Regular Days', 'Rate', 'Regular Pay', 'Ecola', 'Overtime', 'Overtime Pay',
               'Sunday', 'Sunday Pay', 'Sunday OT', 'Sunday OT Pay', 'Sunday ND', 'Sunday ND Amount',
               'Holiday', 'Holiday Pay', 'Holiday OT', 'Holiday OT Pay', 'Special Holiday', 'Special Holiday Pay',
               'Special Holiday OT', 'Special Holiday OT Pay', 'Rest Days', 'Rest Day Pay', 'Rest Day OT', 'Rest Day OT Pay',
               'Night Diff', 'Night Diff Pay', 'Tardiness / Undertime', 'Tardiness / Undertime Deduct.',
               'Uniform', 'Medical', 'Canteen', 'Gatepass', 'Vale', '13th Month', 'SIL', 'TSHIRT', 'RF', 'House',
               'Misc', 'Pants', 'Company Loan', 'SSS Loan', 'PAGIBIG Loan', 'SSS', 'PHILHEALTH', 'PAGIBIG', 'NET', 'Service Fee', 'Employer SSS', 'Employer PAGIBIG', 'Employer PHILHEALTH'
               ]

    for col_num in range(len(columns)):
        ws.write(row_num, col_num, columns[col_num], font_style)

    # Sheet body, remaining rows
#    font_style = xlwt.XFStyle()
    style = 'align: wrap on, vert centre, horiz center'
    font_style = xlwt.Style.easyxf(style)
    rows = Payroll.objects.filter(base_payroll=pk).values_list(
        'employee__last_name', 'regular_days', 'rate', 'regular_amount', 'ecola', 'overtime_regular', 'overtime_regular_amount',
        'sunday', 'sunday_amount', 'sunday_overtime', 'sunday_overtime_amount', 'sunday_nd', 'sunday_nd_amount',
        'holiday_regular_days', 'holiday_regular_amount', 'holiday_overtime', 'holiday_overtime_amount', 'special_holiday_days', 'special_holiday_amount',
        'special_holiday_overtime', 'special_holiday_overtime_amount', 'rest_days', 'rest_amount', 'rest_day_overtime', 'rest_day_overtime_amount',
        'night_diff_days', 'night_diff_amount', 'tardiness_undertime_regular', 'tardiness_undertime_regular_amount',
        'uniform', 'medical', 'canteen', 'gatepass', 'vale', 'thirteenth_month', 'sil', 'tshirt', 'rf', 'house',
        'misc', 'pants', 'company_loan', 'sss_loan', 'pagibig_loan', 'sss', 'philhealth', 'pagibig', 'net_amount', 'service_fee', 'sss_employer', 'pagibig_employer', 'philhealth_employer'
    )
    for row in rows:
        row_num += 1
        for col_num in range(len(row)):
            ws.write(row_num, col_num, row[col_num], font_style)

    # add mo lahat ng dulo
    listed_payroll = Payroll.objects.filter(
        base_payroll=pk)
    print(listed_payroll)
    style = 'align: wrap on, vert centre, horiz center'
    font_style = xlwt.Style.easyxf(style)

    paydict = {
        "regular_days": 0,
        "rate": 0,
        "regular_amount": 0,
        "ecola": 0,
        "overtime_regular": 0,
        "overtime_regular_amount": 0,
        "sunday": 0,
        "sunday_amount": 0,
        "sunday_overtime": 0,
        "sunday_overtime_amount": 0,
        "sunday_nd": 0,
        "sunday_nd_amount": 0,
        "holiday_regular_days": 0,
        "holiday_regular_amount": 0,
        "holiday_overtime": 0,
        "holiday_overtime_amount": 0,
        "special_holiday_days": 0,
        "special_holiday_amount": 0,
        "special_holiday_overtime": 0,
        "special_holiday_overtime_amount": 0,
        "rest_days": 0,
        "rest_amount": 0,
        "rest_day_overtime": 0,
        "rest_day_overtime_amount": 0,
        "night_diff_days": 0,
        "night_diff_amount": 0,
        "tardiness_undertime_regular": 0,
        "tardiness_undertime_regular_amount": 0,
        "uniform": 0,
        "medical": 0,
        "canteen": 0,
        "gatepass": 0,
        "vale": 0,
        "thirteenth_month": 0,
        "sil": 0,
        "tshirt": 0,
        "rf": 0,
        "house": 0,
        "misc": 0,
        "pants": 0,
        "company_loan": 0,
        "sss_loan": 0,
        "pagibig_loan": 0,
        "sss": 0,
        "philhealth": 0,
        "pagibig": 0,
        "net_amount": 0,
        "service_fee": 0,
        "sss_employer": 0,
        "pagibig_employer": 0,
        "philhealth_employer": 0,

    }
    # regular_days = 0
    for lp in listed_payroll:
        paydict["regular_days"] = paydict["regular_days"] + lp.regular_days
        paydict["rate"] = paydict["rate"] + lp.rate
        paydict["regular_amount"] = paydict["regular_amount"] + \
            lp.regular_amount
        paydict["ecola"] = paydict["ecola"] + lp.ecola
        paydict["overtime_regular"] = paydict["overtime_regular"] + \
            lp.overtime_regular
        paydict["overtime_regular_amount"] = paydict["overtime_regular_amount"] + \
            lp.overtime_regular_amount
        paydict["sunday"] = paydict["sunday"] + lp.sunday
        paydict["sunday_amount"] = paydict["sunday_amount"] + lp.sunday_amount
        paydict["sunday_overtime"] = paydict["sunday_overtime"] + \
            lp.sunday_overtime
        paydict["sunday_overtime_amount"] = paydict["sunday_overtime_amount"] + \
            lp.sunday_overtime_amount
        paydict["sunday_nd"] = paydict["sunday_nd"] + lp.sunday_nd
        paydict["sunday_nd_amount"] = paydict["sunday_nd_amount"] + \
            lp.sunday_nd_amount
        paydict["holiday_regular_days"] = paydict["holiday_regular_days"] + \
            lp.holiday_regular_days
        paydict["holiday_regular_amount"] = paydict["holiday_regular_amount"] + \
            lp.holiday_regular_amount
        paydict["holiday_overtime"] = paydict["holiday_overtime"] + \
            lp.holiday_overtime
        paydict["holiday_overtime_amount"] = paydict["holiday_overtime_amount"] + \
            lp.holiday_overtime_amount
        paydict["special_holiday_days"] = paydict["special_holiday_days"] + \
            lp.special_holiday_days
        paydict["special_holiday_amount"] = paydict["special_holiday_amount"] + \
            lp.special_holiday_amount
        paydict["special_holiday_overtime"] = paydict["special_holiday_overtime"] + \
            lp.special_holiday_overtime
        paydict["special_holiday_overtime_amount"] = paydict["special_holiday_overtime_amount"] + \
            lp.special_holiday_overtime_amount
        paydict["rest_days"] = paydict["rest_days"] + lp.rest_days
        paydict["rest_amount"] = paydict["rest_amount"] + lp.rest_amount
        paydict["rest_day_overtime"] = paydict["rest_day_overtime"] + \
            lp.rest_day_overtime
        paydict["rest_day_overtime_amount"] = paydict["rest_day_overtime_amount"] + \
            lp.rest_day_overtime_amount
        paydict["night_diff_days"] = paydict["night_diff_days"] + \
            lp.night_diff_days
        paydict["night_diff_amount"] = paydict["night_diff_amount"] + \
            lp.night_diff_amount
        paydict["tardiness_undertime_regular"] = paydict["tardiness_undertime_regular"] + \
            lp.tardiness_undertime_regular
        paydict["tardiness_undertime_regular_amount"] = paydict["tardiness_undertime_regular_amount"] + \
            lp.tardiness_undertime_regular_amount
        paydict["uniform"] = paydict["uniform"] + lp.uniform
        paydict["medical"] = paydict["medical"] + lp.medical
        paydict["canteen"] = paydict["canteen"] + lp.canteen
        paydict["gatepass"] = paydict["gatepass"] + lp.gatepass
        paydict["vale"] = paydict["vale"] + lp.vale
        paydict["thirteenth_month"] = paydict["thirteenth_month"] + \
            lp.thirteenth_month
        paydict["sil"] = paydict["sil"] + lp.sil
        paydict["tshirt"] = paydict["tshirt"] + lp.tshirt
        paydict["rf"] = paydict["rf"] + lp.rf
        paydict["house"] = paydict["house"] + lp.house
        paydict["misc"] = paydict["misc"] + lp.misc
        paydict["pants"] = paydict["pants"] + lp.pants
        paydict["company_loan"] = paydict["company_loan"] + lp.company_loan
        paydict["sss_loan"] = paydict["sss_loan"] + lp.sss_loan
        paydict["pagibig_loan"] = paydict["pagibig_loan"] + lp.pagibig_loan
        paydict["sss"] = paydict["sss"] + lp.sss
        paydict["philhealth"] = paydict["philhealth"] + lp.philhealth
        paydict["pagibig"] = paydict["pagibig"] + lp.pagibig
        paydict["net_amount"] = paydict["net_amount"] + lp.net_amount
        paydict["service_fee"] = paydict["service_fee"] + lp.service_fee
        paydict["sss_employer"] = paydict["sss_employer"] + lp.sss_employer
        paydict["pagibig_employer"] = paydict["pagibig_employer"] + \
            lp.pagibig_employer
        paydict["philhealth_employer"] = paydict["philhealth_employer"] + \
            lp.philhealth_employer

    print(f"paydict: {listed_payroll}")
    print(f"paydict: {paydict}")

    # for row in rows:
    #     row_num += 1
    #     for col_num in range(len(row)):
    #         ws.write(row_num, col_num, row[col_num], font_style)
    style = 'font: colour red, bold True;align: wrap on, vert centre, horiz center'
    font_style = xlwt.Style.easyxf(style)

    col_for_final = 0
    row_num = row_num + 1
    for final_row in paydict:
        # for col_num in range(46):
        col_for_final = col_for_final + 1

        print(paydict[final_row])
        ws.write(row_num, col_for_final, paydict[final_row], font_style)
    # print(paydict[0])

    wb.save(response)
    return response


@login_required
def payroll_csv_phil_asia(request, pk):
    response = HttpResponse(content_type='text/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="payroll.xls"'

    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('payroll')

    # heading
    # get
    base_payroll = get_object_or_404(Base_payroll, pk=pk)
    print(base_payroll)
    style = 'align: wrap on, horiz center; font: bold on'
#    ws.row(0).write(0, value, xlwt.Style.easyxf(style))
    ws.write_merge(0, 0, 0, 1, 'LUXOR MANPOWER PAYROLL',
                   xlwt.Style.easyxf(style))
    ws.write_merge(1, 1, 0, 1, 'FIRM NAME: PHIL ASIA',
                   xlwt.Style.easyxf(style))
    ws.write_merge(
        2, 2, 0, 3, f'Payroll Covered: {base_payroll}', xlwt.Style.easyxf(style))

    helpers_count = Employee_hiring_details.objects.filter(
        payment_method=base_payroll.payment_method, position="helper").count()
    operators_count = Employee_hiring_details.objects.filter(
        payment_method=base_payroll.payment_method, position="operator").count()
    special_operators_count = Employee_hiring_details.objects.filter(
        payment_method=base_payroll.payment_method, position="special_operator").count()
    drivers_count = Employee_hiring_details.objects.filter(
        payment_method=base_payroll.payment_method, position="driver").count()

    print(helpers_count)
    print(operators_count)
    print(special_operators_count)
    print(drivers_count)
    # get the helpers, operators, special_operators, drivers
    row_num = 4

    style = 'align: wrap on, vert centre, horiz center; font: bold on'
    font_style = xlwt.Style.easyxf(style)
    style_basic = 'align: wrap on, vert centre, horiz center'
    font_style_basic = xlwt.Style.easyxf(style)

    columns = ['DAYS', 'REG. HOURS', 'NIGHT DIFF.',
               'REGULAR OT', 'TOTAL', 'DEDUCTIONS']

    ws.write_merge(
        row_num, 6, 0, 1, f'Helpers', xlwt.Style.easyxf(style))

    # header
    ws.write_merge(
        row_num, row_num, 2, 3, f'DAYS', xlwt.Style.easyxf(style))
    ws.write_merge(
        row_num, row_num, 4, 5, f'REG. HOURS', xlwt.Style.easyxf(style))
    ws.write_merge(
        row_num, row_num, 6, 7, f'NIGHT DIFF.', xlwt.Style.easyxf(style))
    ws.write_merge(
        row_num, row_num, 8, 9, f'REGULAR OT', xlwt.Style.easyxf(style))
    ws.write_merge(
        row_num, row_num, 10, 10, f'TOTAL', xlwt.Style.easyxf(style))
    ws.write_merge(
        row_num, row_num, 12, 15, f'DEDUCTIONS', xlwt.Style.easyxf(style))

    row_num += 1
    # sub header
    ws.write_merge(
        row_num, row_num+1, 2, 2, f'DAYS', xlwt.Style.easyxf(style_basic))
    ws.write_merge(
        row_num, row_num+1, 3, 3, f'AMOUNT', xlwt.Style.easyxf(style_basic))
    ws.write_merge(
        row_num, row_num+1, 4, 4, f'HRS', xlwt.Style.easyxf(style_basic))
    ws.write_merge(
        row_num, row_num+1, 5, 5, f'PAY', xlwt.Style.easyxf(style_basic))
    ws.write_merge(
        row_num, row_num+1, 6, 6, f'ND', xlwt.Style.easyxf(style_basic))
    ws.write_merge(
        row_num, row_num+1, 7, 7, f'ND PAY', xlwt.Style.easyxf(style_basic))
    ws.write_merge(
        row_num, row_num+1, 8, 8, f'HRS', xlwt.Style.easyxf(style_basic))
    ws.write_merge(
        row_num, row_num+1, 9, 9, f'PAY', xlwt.Style.easyxf(style_basic))
    ws.write_merge(
        row_num, row_num+1, 10, 10, f'AMOUNT', xlwt.Style.easyxf(style))
    ws.write_merge(
        row_num, row_num+1, 11, 11, f'CANTEEN', xlwt.Style.easyxf(style))
    ws.write_merge(
        row_num, row_num+1, 12, 12, f'OFFICE', xlwt.Style.easyxf(style))
    ws.write_merge(
        row_num, row_num+1, 13, 13, f'SSS', xlwt.Style.easyxf(style))
    ws.write_merge(
        row_num, row_num+1, 14, 14, f'PHILHEALTH', xlwt.Style.easyxf(style))
    ws.write_merge(
        row_num, row_num+1, 15, 15, f'PAGIBIG', xlwt.Style.easyxf(style))
    ws.write_merge(
        row_num, row_num+1, 16, 16, f'NET AMOUNT', xlwt.Style.easyxf(style))
    ws.write_merge(
        row_num, row_num+1, 17, 17, f'PAYEE SIGNATURE', xlwt.Style.easyxf(style))

    # paydict = {
    #     "regular_days": 0,

    # }

    listed_payroll = Payroll_for_phil_asia.objects.filter(base_payroll=pk)

    may_operator_na = 0
    may_special_operator_na = 0
    may_driver_na = 0

    print(row_num)
    # loop helpers
    col_for_final = 0
    row_num += 2
    for helpers in listed_payroll:
        # col_for_final += 1
        # print(helpers.em)(
        if helpers.employee.get_hiring_details().position == "helper":
            ws.write_merge(row_num, row_num, col_for_final, col_for_final+1,
                           f'{helpers.employee.first_name}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+2, col_for_final+2,
                           f'{helpers.days}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+3, col_for_final+3,
                           f'{helpers.amount}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+4, col_for_final +
                           4, f'{helpers.hrs}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+5, col_for_final +
                           5, f'{helpers.pay}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+6, col_for_final +
                           6, f'{helpers.nd}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+7, col_for_final+7,
                           f'{helpers.nd_pay}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+8, col_for_final+8,
                           f'{helpers.hrs_ot}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+9, col_for_final+9,
                           f'{helpers.hrs_ot_pay}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+10, col_for_final+10,
                           f'{helpers.total_amount}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+11, col_for_final+11,
                           f'{helpers.canteen}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+12, col_for_final +
                           12, f'{helpers.office}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+13, col_for_final +
                           13, f'{helpers.sss}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+14, col_for_final+14,
                           f'{helpers.philhealth}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+15, col_for_final+15,
                           f'{helpers.pagibig}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+16, col_for_final+16,
                           f'{helpers.net_amount}', xlwt.Style.easyxf(style_basic))

        # ws.write(row_num, col_for_final, helpers.employee.first_name, font_style)
        # ws.write(row_num, col_for_final+1, helpers.employee.first_name, font_style)
            row_num += 1
        if helpers.employee.get_hiring_details().position == "operator":

            if may_operator_na == 0:
                ws.write_merge(row_num, row_num, 0, 1,
                               f'Operators', xlwt.Style.easyxf(style))
                row_num += 1
                may_operator_na = 1

            ws.write_merge(row_num, row_num, col_for_final, col_for_final+1,
                           f'{helpers.employee.first_name}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+2, col_for_final+2,
                           f'{helpers.days}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+3, col_for_final+3,
                           f'{helpers.amount}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+4, col_for_final +
                           4, f'{helpers.hrs}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+5, col_for_final +
                           5, f'{helpers.pay}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+6, col_for_final +
                           6, f'{helpers.nd}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+7, col_for_final+7,
                           f'{helpers.nd_pay}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+8, col_for_final+8,
                           f'{helpers.hrs_ot}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+9, col_for_final+9,
                           f'{helpers.hrs_ot_pay}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+10, col_for_final+10,
                           f'{helpers.total_amount}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+11, col_for_final+11,
                           f'{helpers.canteen}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+12, col_for_final +
                           12, f'{helpers.office}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+13, col_for_final +
                           13, f'{helpers.sss}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+14, col_for_final+14,
                           f'{helpers.philhealth}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+15, col_for_final+15,
                           f'{helpers.pagibig}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+16, col_for_final+16,
                           f'{helpers.net_amount}', xlwt.Style.easyxf(style_basic))
            row_num += 1

        if helpers.employee.get_hiring_details().position == "special_operator":

            if may_special_operator_na == 0:
                ws.write_merge(row_num, row_num, 0, 1,
                               f'Special Operators', xlwt.Style.easyxf(style))
                row_num += 1
                may_special_operator_na = 1
            ws.write_merge(row_num, row_num, col_for_final, col_for_final+1,
                           f'{helpers.employee.first_name}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+2, col_for_final+2,
                           f'{helpers.days}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+3, col_for_final+3,
                           f'{helpers.amount}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+4, col_for_final +
                           4, f'{helpers.hrs}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+5, col_for_final +
                           5, f'{helpers.pay}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+6, col_for_final +
                           6, f'{helpers.nd}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+7, col_for_final+7,
                           f'{helpers.nd_pay}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+8, col_for_final+8,
                           f'{helpers.hrs_ot}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+9, col_for_final+9,
                           f'{helpers.hrs_ot_pay}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+10, col_for_final+10,
                           f'{helpers.total_amount}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+11, col_for_final+11,
                           f'{helpers.canteen}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+12, col_for_final +
                           12, f'{helpers.office}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+13, col_for_final +
                           13, f'{helpers.sss}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+14, col_for_final+14,
                           f'{helpers.philhealth}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+15, col_for_final+15,
                           f'{helpers.pagibig}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+16, col_for_final+16,
                           f'{helpers.net_amount}', xlwt.Style.easyxf(style_basic))
            row_num += 1

        if helpers.employee.get_hiring_details().position == "driver":

            if may_driver_na == 0:
                ws.write_merge(row_num, row_num, 0, 1,
                               f'Drivers', xlwt.Style.easyxf(style))
                row_num += 1
                may_driver_na = 1

            ws.write_merge(row_num, row_num, col_for_final, col_for_final+1,
                           f'{helpers.employee.first_name}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+2, col_for_final+2,
                           f'{helpers.days}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+3, col_for_final+3,
                           f'{helpers.amount}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+4, col_for_final +
                           4, f'{helpers.hrs}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+5, col_for_final +
                           5, f'{helpers.pay}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+6, col_for_final +
                           6, f'{helpers.nd}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+7, col_for_final+7,
                           f'{helpers.nd_pay}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+8, col_for_final+8,
                           f'{helpers.hrs_ot}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+9, col_for_final+9,
                           f'{helpers.hrs_ot_pay}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+10, col_for_final+10,
                           f'{helpers.total_amount}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+11, col_for_final+11,
                           f'{helpers.canteen}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+12, col_for_final +
                           12, f'{helpers.office}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+13, col_for_final +
                           13, f'{helpers.sss}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+14, col_for_final+14,
                           f'{helpers.philhealth}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+15, col_for_final+15,
                           f'{helpers.pagibig}', xlwt.Style.easyxf(style_basic))
            ws.write_merge(row_num, row_num, col_for_final+16, col_for_final+16,
                           f'{helpers.net_amount}', xlwt.Style.easyxf(style_basic))
            row_num += 1

    # for col_num in range(len(columns)):
    #     # ws.write_merge(row_num, row_num, col_num+2, col_num1, f'Helpers', xlwt.Style.easyxf(style))
    #     ws.write(row_num, col_for_final, paydict[final_row], font_style)

    wb.save(response)
    return response


@login_required
def payroll_csv(request, pk):
    response = HttpResponse(content_type='text/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="payroll.xls"'

    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('payroll')

    # heading
    # get
    base_payroll = get_object_or_404(Base_payroll, pk=pk)
    print(base_payroll)
    style = 'align: wrap on, vert centre, horiz center; font: bold on'
#    ws.row(0).write(0, value, xlwt.Style.easyxf(style))
    ws.write_merge(0, 1, 0, 20, 'Luxor Manpower Payroll',
                   xlwt.Style.easyxf(style))
    ws.write_merge(
        2, 2, 0, 4, f'Payroll Covered: {base_payroll}', xlwt.Style.easyxf(style))
    # ws.write_merge(1, 0, 0, 0, 'Payroll Covered', xlwt.Style.easyxf(style))
    # Sheet header, first row

    row_num = 4

    style = 'align: wrap on, vert centre, horiz center; font: bold on'
    font_style = xlwt.Style.easyxf(style)
    columns = ['Employee', 'Regular Days', 'Rate', 'Regular Pay', 'Ecola', 'Overtime', 'Overtime Pay',
               'Sunday', 'Sunday Pay', 'Sunday OT', 'Sunday OT Pay', 'Sunday ND', 'Sunday ND Amount',
               'Holiday', 'Holiday Pay', 'Holiday OT', 'Holiday OT Pay', 'Special Holiday', 'Special Holiday Pay',
               'Special Holiday OT', 'Special Holiday OT Pay', 'Rest Days', 'Rest Day Pay', 'Rest Day OT', 'Rest Day OT Pay',
               'Night Diff', 'Night Diff Pay', 'Tardiness / Undertime', 'Tardiness / Undertime Deduct.',
               'Uniform', 'Medical', 'Canteen', 'Gatepass', 'Vale', '13th Month', 'SIL', 'TSHIRT', 'RF', 'House',
               'Misc', 'Pants', 'Company Loan', 'SSS Loan', 'PAGIBIG Loan', 'SSS', 'PHILHEALTH', 'PAGIBIG', 'NET'
               ]

    for col_num in range(len(columns)):
        ws.write(row_num, col_num, columns[col_num], font_style)

    # Sheet body, remaining rows
#    font_style = xlwt.XFStyle()
    style = 'align: wrap on, vert centre, horiz center'
    font_style = xlwt.Style.easyxf(style)
    rows = Payroll.objects.filter(base_payroll=pk).values_list(
        'employee__last_name', 'regular_days', 'rate', 'regular_amount', 'ecola', 'overtime_regular', 'overtime_regular_amount',
        'sunday', 'sunday_amount', 'sunday_overtime', 'sunday_overtime_amount', 'sunday_nd', 'sunday_nd_amount',
        'holiday_regular_days', 'holiday_regular_amount', 'holiday_overtime', 'holiday_overtime_amount', 'special_holiday_days', 'special_holiday_amount',
        'special_holiday_overtime', 'special_holiday_overtime_amount', 'rest_days', 'rest_amount', 'rest_day_overtime', 'rest_day_overtime_amount',
        'night_diff_days', 'night_diff_amount', 'tardiness_undertime_regular', 'tardiness_undertime_regular_amount',
        'uniform', 'medical', 'canteen', 'gatepass', 'vale', 'thirteenth_month', 'sil', 'tshirt', 'rf', 'house',
        'misc', 'pants', 'company_loan', 'sss_loan', 'pagibig_loan', 'sss', 'philhealth', 'pagibig', 'net_amount'
    )
    for row in rows:
        row_num += 1
        for col_num in range(len(row)):
            ws.write(row_num, col_num, row[col_num], font_style)

    # add mo lahat ng dulo
    listed_payroll = Payroll.objects.filter(
        base_payroll=pk)
    print(listed_payroll)
    style = 'align: wrap on, vert centre, horiz center'
    font_style = xlwt.Style.easyxf(style)

    paydict = {
        "regular_days": 0,
        "rate": 0,
        "regular_amount": 0,
        "ecola": 0,
        "overtime_regular": 0,
        "overtime_regular_amount": 0,
        "sunday": 0,
        "sunday_amount": 0,
        "sunday_overtime": 0,
        "sunday_overtime_amount": 0,
        "sunday_nd": 0,
        "sunday_nd_amount": 0,
        "holiday_regular_days": 0,
        "holiday_regular_amount": 0,
        "holiday_overtime": 0,
        "holiday_overtime_amount": 0,
        "special_holiday_days": 0,
        "special_holiday_amount": 0,
        "special_holiday_overtime": 0,
        "special_holiday_overtime_amount": 0,
        "rest_days": 0,
        "rest_amount": 0,
        "rest_day_overtime": 0,
        "rest_day_overtime_amount": 0,
        "night_diff_days": 0,
        "night_diff_amount": 0,
        "tardiness_undertime_regular": 0,
        "tardiness_undertime_regular_amount": 0,
        "uniform": 0,
        "medical": 0,
        "canteen": 0,
        "gatepass": 0,
        "vale": 0,
        "thirteenth_month": 0,
        "sil": 0,
        "tshirt": 0,
        "rf": 0,
        "house": 0,
        "misc": 0,
        "pants": 0,
        # "service_fee": 0,
        "company_loan": 0,
        "sss_loan": 0,
        "pagibig_loan": 0,
        "sss": 0,
        "philhealth": 0,
        "pagibig": 0,
        "net_amount": 0,
    }
    # regular_days = 0
    for lp in listed_payroll:
        paydict["regular_days"] = paydict["regular_days"] + lp.regular_days
        paydict["rate"] = paydict["rate"] + lp.rate
        paydict["regular_amount"] = paydict["regular_amount"] + \
            lp.regular_amount
        paydict["ecola"] = paydict["ecola"] + lp.ecola
        paydict["overtime_regular"] = paydict["overtime_regular"] + \
            lp.overtime_regular
        paydict["overtime_regular_amount"] = paydict["overtime_regular_amount"] + \
            lp.overtime_regular_amount
        paydict["sunday"] = paydict["sunday"] + lp.sunday
        paydict["sunday_amount"] = paydict["sunday_amount"] + lp.sunday_amount
        paydict["sunday_overtime"] = paydict["sunday_overtime"] + \
            lp.sunday_overtime
        paydict["sunday_overtime_amount"] = paydict["sunday_overtime_amount"] + \
            lp.sunday_overtime_amount
        paydict["sunday_nd"] = paydict["sunday_nd"] + lp.sunday_nd
        paydict["sunday_nd_amount"] = paydict["sunday_nd_amount"] + \
            lp.sunday_nd_amount
        paydict["holiday_regular_days"] = paydict["holiday_regular_days"] + \
            lp.holiday_regular_days
        paydict["holiday_regular_amount"] = paydict["holiday_regular_amount"] + \
            lp.holiday_regular_amount
        paydict["holiday_overtime"] = paydict["holiday_overtime"] + \
            lp.holiday_overtime
        paydict["holiday_overtime_amount"] = paydict["holiday_overtime_amount"] + \
            lp.holiday_overtime_amount
        paydict["special_holiday_days"] = paydict["special_holiday_days"] + \
            lp.special_holiday_days
        paydict["special_holiday_amount"] = paydict["special_holiday_amount"] + \
            lp.special_holiday_amount
        paydict["special_holiday_overtime"] = paydict["special_holiday_overtime"] + \
            lp.special_holiday_overtime
        paydict["special_holiday_overtime_amount"] = paydict["special_holiday_overtime_amount"] + \
            lp.special_holiday_overtime_amount
        paydict["rest_days"] = paydict["rest_days"] + lp.rest_days
        paydict["rest_amount"] = paydict["rest_amount"] + lp.rest_amount
        paydict["rest_day_overtime"] = paydict["rest_day_overtime"] + \
            lp.rest_day_overtime
        paydict["rest_day_overtime_amount"] = paydict["rest_day_overtime_amount"] + \
            lp.rest_day_overtime_amount
        paydict["night_diff_days"] = paydict["night_diff_days"] + \
            lp.night_diff_days
        paydict["night_diff_amount"] = paydict["night_diff_amount"] + \
            lp.night_diff_amount
        paydict["tardiness_undertime_regular"] = paydict["tardiness_undertime_regular"] + \
            lp.tardiness_undertime_regular
        paydict["tardiness_undertime_regular_amount"] = paydict["tardiness_undertime_regular_amount"] + \
            lp.tardiness_undertime_regular_amount
        paydict["uniform"] = paydict["uniform"] + lp.uniform
        paydict["medical"] = paydict["medical"] + lp.medical
        paydict["canteen"] = paydict["canteen"] + lp.canteen
        paydict["gatepass"] = paydict["gatepass"] + lp.gatepass
        paydict["vale"] = paydict["vale"] + lp.vale
        paydict["thirteenth_month"] = paydict["thirteenth_month"] + \
            lp.thirteenth_month
        paydict["sil"] = paydict["sil"] + lp.sil
        paydict["tshirt"] = paydict["tshirt"] + lp.tshirt
        paydict["rf"] = paydict["rf"] + lp.rf
        paydict["house"] = paydict["house"] + lp.house
        paydict["misc"] = paydict["misc"] + lp.misc
        paydict["pants"] = paydict["pants"] + lp.pants
        # paydict["service_fee"] = paydict["service_fee"] + lp.service_fee
        paydict["company_loan"] = paydict["company_loan"] + lp.company_loan
        paydict["sss_loan"] = paydict["sss_loan"] + lp.sss_loan
        paydict["pagibig_loan"] = paydict["pagibig_loan"] + lp.pagibig_loan
        paydict["sss"] = paydict["sss"] + lp.sss
        paydict["philhealth"] = paydict["philhealth"] + lp.philhealth
        paydict["pagibig"] = paydict["pagibig"] + lp.pagibig
        paydict["net_amount"] = paydict["net_amount"] + lp.net_amount
    print(f"paydict: {listed_payroll}")
    print(f"paydict: {paydict}")

    # for row in rows:
    #     row_num += 1
    #     for col_num in range(len(row)):
    #         ws.write(row_num, col_num, row[col_num], font_style)
    style = 'font: colour red, bold True;align: wrap on, vert centre, horiz center'
    font_style = xlwt.Style.easyxf(style)

    col_for_final = 0
    row_num = row_num + 1
    for final_row in paydict:
        # for col_num in range(46):
        col_for_final = col_for_final + 1

        print(paydict[final_row])
        ws.write(row_num, col_for_final, paydict[final_row], font_style)
    # print(paydict[0])
    wb.save(response)
    return response


def go_with_the_phil_asia(found, company_to, saved_form):

    emp_rate = found.rate
    training_rate = found.training_rate

    if found.employee.gov_deductions_to_implement == "company_base_deductions":
        sss = company_to.sss
        pagibig = company_to.pagibig
        philhealth = company_to.philhealth
    elif found.employee.gov_deductions_to_implement == "employee_base_deductions":
        # to edit for bracketing or manual
        # check sss if manual ba or bracket
        pagibig = found.employee.pagibig_value
        philhealth = found.employee.philhealth_value
        if found.employee.sss_option == "manual":
            sss = found.employee.sss_value
        elif found.employee.sss_option == "bracket":
            sss = found.employee.sss_bracket

    Payroll_for_phil_asia.objects.create(
        base_payroll=saved_form,
        employee=found.employee,
        rate=emp_rate,
        training_rate=training_rate,

        # sss=sss,
        # pagibig=pagibig,
        # philhealth=philhealth
    )


# payroll list

@login_required
def payroll_list(request):
    if request.method == 'POST':
        form = BasePayrollForm(request.POST)
        if form.is_valid():

            # check if existing na yung base payroll
            # return HttpResponse(form)
            existing_na = Base_payroll.objects.filter(company=form.cleaned_data['company'], start_date=form.cleaned_data['start_date'], end_date=form.cleaned_data['end_date'], payment_method=form.cleaned_data['payment_method'])
            if existing_na:
                messages.error(request, "Payroll already exists.")
                return redirect('payroll-list')

            # check if payroll is phil-asia
            form_company = str(form.cleaned_data['company'])
            if "phil asia" in form_company:
                # return HttpResponse("phil asia to")
                    # check if meron ng existings
                meron_ng_payroll = Base_payroll.objects.filter(company=form.cleaned_data['company'], start_date=form.cleaned_data[
                                                               'start_date'], end_date=form.cleaned_data['end_date'], payment_method=form.cleaned_data['payment_method'])
                if meron_ng_payroll:
                    messages.error(request, f'Payroll already exists.')
                    return redirect('payroll-list')
                saved_form = form.save()

                # get employee with the company, payment method, per position
                found_helper_employees = Employee_hiring_details.objects.filter(position="helper", payment_method=saved_form.payment_method, employee__company=saved_form.company,
                                                                                employee__contract_expiration__gte=datetime.now()-timedelta(days=5)
                                                                                )
                found_operator_employees = Employee_hiring_details.objects.filter(position="operator", payment_method=saved_form.payment_method, employee__company=saved_form.company,
                                                                                  employee__contract_expiration__gte=datetime.now()-timedelta(days=5)
                                                                                  )
                found_special_operator_employees = Employee_hiring_details.objects.filter(position="special_operator", payment_method=saved_form.payment_method, employee__company=saved_form.company,
                                                                                          employee__contract_expiration__gte=datetime.now()-timedelta(days=5)
                                                                                          )
                found_driver_employees = Employee_hiring_details.objects.filter(position="driver", payment_method=saved_form.payment_method, employee__company=saved_form.company,
                                                                                employee__contract_expiration__gte=datetime.now()-timedelta(days=5)
                                                                                )

                company_to = get_object_or_404(
                    Company_rates, company=saved_form.company)

                for founds in found_helper_employees:
                    go_with_the_phil_asia(founds, company_to, saved_form)

                for founds in found_operator_employees:
                    go_with_the_phil_asia(founds, company_to, saved_form)

                for founds in found_special_operator_employees:
                    go_with_the_phil_asia(founds, company_to, saved_form)

                for founds in found_driver_employees:
                    go_with_the_phil_asia(founds, company_to, saved_form)

                messages.success(request, f'Payroll was successfully created')
                return redirect('payroll-list')

            else:
                meron_ng_payroll = Base_payroll.objects.filter(company=form.cleaned_data['company'], start_date=form.cleaned_data[
                                                               'start_date'], end_date=form.cleaned_data['end_date'], payment_method=form.cleaned_data['payment_method'])
                if meron_ng_payroll:
                    messages.error(request, f'Payroll already exists.')
                    return redirect('payroll-list')
                # save to base payroll
                saved_form = form.save()
                # get employee with the company, payment method
                found_employees = Employee_hiring_details.objects.filter(payment_method=saved_form.payment_method, employee__company=saved_form.company,
                                                                         employee__contract_expiration__gte=datetime.now()-timedelta(days=5)
                                                                         )
                company_to = get_object_or_404(
                    Company_rates, company=saved_form.company)
                # return HttpResponse(company_to.base_rate)
                # save found employees to Payroll
                for found in found_employees:

                    # get pagibig loan
                    pagibig_loan = 0.0
                    if saved_form.activate_gov_deductions:
                        found_pagibig_loan = Employee_pagibig_loan.objects.filter(
                            employee=found.employee)
                        print(found_pagibig_loan)
                        for pagibig in found_pagibig_loan:
                            if pagibig.status == False:
                                pagibig_loan = pagibig.rate_to_deduct
                                print("pagibig loan to")
                                print(pagibig_loan)


                    # get sss loan
                    # check if activate_gov_deductions
                    sss_loan = 0.0
                    if saved_form.activate_gov_deductions:
                        found_sss_loan = Employee_sss_loan.objects.filter(
                            employee=found.employee)
                        print(found_sss_loan)
                        for sss in found_sss_loan:
                            if sss.status == False:
                                sss_loan = sss.rate_to_deduct
                                print("sss loan to")
                                print(sss_loan)

                    # get vale
                    found_vale = Employee_vale.objects.filter(
                        employee=found.employee)
                    vale = 0.0
                    print(found_vale)
                    for val in found_vale:
                        if val.status == False:
                            vale = val.rate_to_deduct

                    # get company loan
                    found_company_loans = Employee_company_loan.objects.filter(
                        employee=found.employee)
                    company_loan = 0.0
                    for comloan in found_company_loans:
                        if comloan.status == False:
                            company_loan = comloan.rate_to_deduct
                            # # save to Employee_comloan_contrib
                            # Employee_comloan_contrib.create(
                            #   employee = found.employee,
                            #   cut_off_date = saved_form.end_date,
                            #   company_loan = comloan.id,
                            #   contribution_collected = company_loan
                            # )

                    # get canteen
                    found_canteen_loans = Employee_canteen.objects.filter(
                        employee=found.employee)
                    canteen_loan = 0.0
                    for canteen in found_canteen_loans:
                        if canteen.status == False:
                            canteen_loan = canteen.rate_to_deduct

                    # get medical
                    found_medical_loans = Employee_medical.objects.filter(
                        employee=found.employee)
                    medical_loan = 0.0
                    for medical in found_medical_loans:
                        if medical.status == False:
                            medical_loan = medical.rate_to_deduct

                    # get gatepass
                    found_gatepass_loans = Employee_gatepass.objects.filter(
                        employee=found.employee)
                    gatepass_loan = 0.0
                    for gatepass in found_gatepass_loans:
                        if gatepass.status == False:
                            gatepass_loan = gatepass.rate_to_deduct


                    emp_rate = found.rate
                    training_rate = found.training_rate

                    # get government deductions from company

                    # check company base or employee base

                    if found.employee.gov_deductions_to_implement == "company_base_deductions":
                        sss = company_to.sss
                        pagibig = company_to.pagibig
                        philhealth = company_to.philhealth
                    elif found.employee.gov_deductions_to_implement == "employee_base_deductions":
                        # to edit for bracketing or manual
                        # check sss if manual ba or bracket
                        pagibig = found.employee.pagibig_value
                        philhealth = found.employee.philhealth_value
                        if found.employee.sss_option == "manual":
                            sss = found.employee.sss_value
                        elif found.employee.sss_option == "bracket":
                            sss = found.employee.sss_bracket

                    Payroll.objects.create(
                        base_payroll=saved_form,
                        employee=found.employee,
                        rate=emp_rate,
                        training_rate=training_rate,
                        company_loan=company_loan,
                        gatepass=gatepass_loan,
                        canteen=canteen_loan,
                        medical=medical_loan,
                        vale=vale,
                        sss_loan=sss_loan,
                        pagibig_loan=pagibig_loan,
                        sss=sss,
                        pagibig=pagibig,
                        philhealth=philhealth
                    )
                action = f"New payroll was successfully added. (company={form.cleaned_data['company']}, start_date={form.cleaned_data['start_date']}, end_date={form.cleaned_data['end_date']}, payment_method={form.cleaned_data['payment_method']} activate government deductions = {form.cleaned_data['activate_gov_deductions']})"
                Logs.objects.create(
                    action=action, action_by=request.user, action_date=datetime.now())
                messages.success(request, f'Payroll was successfully created')
                return redirect('payroll-list')

        messages.error(
            request, f'Certain fields are required. Click the button for more info.')
    else:
        form = BasePayrollForm()

    may_payroll_list = Base_payroll.objects.all()
    # return print(may_payroll_list)
    context = {
        'title': 'company',
        'head': 'Payroll',
        'form': form,
        'payroll_list_to': 1,
        'may_payroll_list': may_payroll_list
    }

    return render(request, 'payrolllist/payroll_list.html', context)

# delete base payroll


@login_required
def payroll_delete_base(request, pk):
    Base_payroll.objects.filter(pk=pk).delete()
    messages.success(request, 'Payroll was successfully deleted')
    return redirect('payroll-list')


@login_required  # encode
def payroll_encode(request, payroll_view, pk):
    base_payroll = get_object_or_404(Base_payroll, pk=payroll_view)
    employee = get_object_or_404(Payroll, pk=pk)
    company_rates = Company_rates.objects.filter(
        company=base_payroll.company).first()

    if request.method == 'POST':
        # form_ecola = PayrollEcolaForm(request.POST or None, instance=base_payroll)
        form_fulladd = PayrollFullAddForm(
            request.POST or None, instance=employee)
        # form_holiday = PayrollHolidayForm(request.POST or None, instance=base_payroll)
        # form_overtime = PayrollOvertimeForm(request.POST or None, instance=base_payroll)
        # form_companydeduction = PayrollCompanyDeductionsForm(request.POST or None, instance=base_payroll)
        # form_governmentdeductions = PayrollGovernmentDeductionsForm(request.POST or None, instance=base_payroll)
        # form_restday = PayrollRestDayForm(request.POST or None, instance=base_payroll)
        # form_income = PayrollIncomeForm(request.POST or None, instance=base_payroll)
        # form_companydeduction2 = PayrollCompanyDeductions2Form(request.POST or None, instance=base_payroll)

        # return HttpResponse(request.POST.get('overtime_regular'))
        # return HttpResponse(request.POST.get('overtime_regular_amount'))
        # if form_ecola.is_valid() and form_fulladd.is_valid() and form_holiday.is_valid() and form_overtime.is_valid() and form_companydeduction.is_valid() and form_governmentdeductions.is_valid() and form_restday.is_valid() and form_income.is_valid() and form_companydeduction2.is_valid():
        if form_fulladd.is_valid():

            form_fulladd.save()
            # form_ecola.save()
            messages.success(
                request, f'Employee {employee} was successfully updated')
            return redirect('payroll-view', pk=payroll_view)
        # else:
        #     return HttpResponse('erro daw');

    # variables

    form = PayrollFullAddForm(instance=employee)
    # holiday_form = PayrollHolidayForm()
    # overtime_form = PayrollOvertimeForm()
    # companydeduction_form = PayrollCompanyDeductionsForm()
    # governmentdeductions_form = PayrollGovernmentDeductionsForm(instance=employee)
    # restday_form = PayrollRestDayForm()
    # income_form = PayrollIncomeForm()
    # companydeduction2_form = PayrollCompanyDeductions2Form()

    context = {
        'title': 'company',
        'head': f'Payroll Encode - {employee.employee}',
        'form': form,
        # 'holiday_form': holiday_form,
        # 'overtime_form': overtime_form,
        # 'restday_form': restday_form,
        # 'companydeduction_form': companydeduction_form,
        # 'governmentdeductions_form': governmentdeductions_form,
        # 'companydeduction2_form': companydeduction2_form,
        # 'income_form': income_form,
        # 'ecola_form': ecola_form,
        'payroll_view': payroll_view,
        'base_payroll': base_payroll,
        'employee_name': employee.employee
    }
    return render(request, 'payrolllist/payroll_encode.html', context)

# payroll view ni phil Asia


def payroll_view_phil_asia(request, pk):
    found_emp_payroll = Payroll_for_phil_asia.objects.filter(base_payroll=pk)

    base_payroll = get_object_or_404(Base_payroll, pk=pk)

    if request.method == "POST":
        form = PayrollFullAddFormPhilAsia(request.POST)
        if form.is_valid():
            pays = form.save()
            Logs.objects.create(
                action=f"Payroll for {base_payroll} was sucessfully updated", action_by=request.user, action_date=datetime.now())
            messages.success(request, f'Payroll was successfully saved')
            return redirect('payroll-view-phil-asia', pk=pk)

    # return HttpResponse(f"base to {pk} - {found_emp_payroll.count()}")
    formset = PayrollFullAddFormPhilAsia(queryset=found_emp_payroll)
    context = {
        'title': 'company',
        'head': f'Payroll View - ({base_payroll})',
        'found_emp_payroll': found_emp_payroll,
        'formset': formset,
        'payroll_view_id': pk,
        'base_payroll': base_payroll
    }

    return render(request, 'payrolllist/payroll_list_phil_asia.html', context)

# payroll view
@login_required
def payroll_view(request, pk):
    # for employee expiration
    # found_emp_payroll = Payroll.objects.filter(base_payroll=pk, employee__contract_expiration__gte=datetime.now()-timedelta(days=5))
    found_emp_payroll = Payroll.objects.filter(base_payroll=pk)

    base_payroll = get_object_or_404(Base_payroll, pk=pk)
    all_employees = Employee_hiring_details.objects.values("employee_id", "overtime_formula", "employee_id__first_name")
    # .values_list(
            # 'employee__pagibig', 'employee__emp_id', 'employee__last_name', 'employee__first_name', 'employee__middle_name', 'pagibig', 'pagibig_employer',
            # 'employee__tin', 'employee__date_of_birth')
    if request.method == "POST":
        form = PayrollFullAddForm(request.POST)
        # return HttpResponse(form)
        if form.is_valid():
            pays = form.save(commit=False)

            for pay in pays:
                # return HttpResponse(pay.employee.id)
                # return HttpResponse(pay.valid_for_deduct_company_loan)
                # check if not valid for deduct company loan, change to zero company loan if not valid
                if pay.valid_for_deduct_company_loan == 0:

                    pay.company_loan = 0

                if pay.valid_for_deduct_vale == 0:
                    pay.vale = 0

                if pay.valid_for_deduct_canteen == 0:
                    pay.canteen = 0

                if pay.valid_for_deduct_medical == 0:
                    pay.medical = 0

                if pay.valid_for_deduct_gatepass == 0:
                    pay.gatepass = 0
                if pay.valid_for_deduct_sss_loan == 0:
                    pay.sss_loan = 0
                if pay.valid_for_deduct_pagibig_loan == 0:
                    pay.pagibig_loan = 0

                saved_payroll = pay.save()
                # print(f"13 - {pay.thirtheenth_month}")

                # check existing company loan if meron
                found_company_loans = Employee_company_loan.objects.filter(
                    employee=pay.employee, status=False)
                # print(f'payroll id+ {pay.id}')

                print(f'payroll id+ {pay.employee}')

                # gawin lang to kapag may loan
                if pay.company_loan > 0:
                    for comloan in found_company_loans:
                        if comloan.status == False:

                            company_loan = comloan.rate_to_deduct
                            # find com loan contrib by payroll id
                            found_comloan_contrib = Employee_comloan_contrib.objects.filter(
                                payroll=pay.id).count()

                            if found_comloan_contrib > 0:
                                # update
                                found_comloan_contrib_got = Employee_comloan_contrib.objects.get(
                                    payroll=pay.id, employee=pay.employee)

                                found_comloan_contrib_got.contribution_collected = company_loan
                                found_comloan_contrib_got.save()

                                # check if paid na
                                total_contributed = Employee_comloan_contrib.objects.filter(
                                    company_loan=comloan.id, employee=pay.employee)
                                # return HttpResponse(total_contributed)
                                total_conts = 0
                                for tots in total_contributed:
                                    total_conts = total_conts + tots.contribution_collected

                                if total_conts >= comloan.load_amount:
                                    # save status to true
                                    comloanloan = get_object_or_404(
                                        Employee_company_loan, employee=pay.employee, status=False)
                                    comloanloan.status = True
                                    comloanloan.save()
                                print("saved..")
                            else:
                                # add / insert

                                # Employee_comloan_contrib
                                # # save to Employee_comloan_contrib
                                Employee_comloan_contrib.objects.create(
                                    employee=pay.employee,
                                    cut_off_date=base_payroll.end_date,
                                    company_loan=comloan,
                                    contribution_collected=company_loan,
                                    payroll=pay.id
                                )

                                # check if paid na
                                total_contributed = Employee_comloan_contrib.objects.filter(
                                    company_loan=comloan.id, employee=pay.employee)
                                # return HttpResponse(total_contributed)
                                total_conts = 0
                                for tots in total_contributed:
                                    total_conts = total_conts + tots.contribution_collected

                                if total_conts >= comloan.load_amount:
                                    # save status to true
                                    comloanloan = get_object_or_404(
                                        Employee_company_loan, employee=pay.employee, status=False)
                                    comloanloan.status = True
                                    comloanloan.save()

                        else:
                            print("sfsdf")

                # check existing sss loan if meron

                if base_payroll.activate_gov_deductions:
                    found_sss_loans = Employee_sss_loan.objects.filter(
                        employee=pay.employee, status=False)
                    print(f'payroll id+ {pay.id}')

                    if pay.sss_loan > 0:
                        for sssloan in found_sss_loans:
                            if sssloan.status == False:
                                sss_loan = sssloan.rate_to_deduct
                                # find sss loan contrib by payroll id
                                found_sssloan_contrib = Employee_sssloan_contrib.objects.filter(
                                    payroll=pay.id).count()
                                if found_sssloan_contrib > 0:
                                    # update

                                    found_sssloan_contrib_got = Employee_sssloan_contrib.objects.get(
                                        payroll=pay.id)
                                    found_sssloan_contrib_got.contribution_collected = sss_loan
                                    found_sssloan_contrib_got.save()

                                    # check if paid na
                                    total_contributed = Employee_sssloan_contrib.objects.filter(
                                        sss_loan=sssloan.id, employee=pay.employee)
                                    # return HttpResponse(total_contributed)
                                    total_conts = 0
                                    for tots in total_contributed:
                                        total_conts = total_conts + tots.contribution_collected

                                    if total_conts >= sssloan.load_amount:
                                        # save status to true
                                        sssloan = get_object_or_404(
                                            Employee_sss_loan, employee=pay.employee, status=False)
                                        sssloan.status = True
                                        sssloan.save()

                                    print("saved..")
                                else:
                                    # add / insert

                                    # Employee_comloan_contrib
                                    # # save to Employee_comloan_contrib
                                    Employee_sssloan_contrib.objects.create(
                                        employee=pay.employee,
                                        cut_off_date=base_payroll.end_date,
                                        sss_loan=sssloan,
                                        contribution_collected=sss_loan,
                                        payroll=pay.id
                                    )

                                    # check if paid na
                                    total_contributed = Employee_sssloan_contrib.objects.filter(
                                        sss_loan=sssloan.id, employee=pay.employee)
                                    # return HttpResponse(total_contributed)
                                    total_conts = 0
                                    for tots in total_contributed:
                                        total_conts = total_conts + tots.contribution_collected

                                    if total_conts >= sssloan.load_amount:
                                        # save status to true
                                        sssloan = get_object_or_404(
                                            Employee_sss_loan, employee=pay.employee, status=False)
                                        sssloan.status = True
                                        sssloan.save()
                            else:
                                print("sfsdf")

                    # check existing pagibig loan if meron
                    found_pagibig_loans = Employee_pagibig_loan.objects.filter(
                        employee=pay.employee, status=False)
                    print(f'payroll id+ {pay.id}')

                    if pay.pagibig_loan > 0:
                        for pagibigloan in found_pagibig_loans:
                            if pagibigloan.status == False:
                                pagibig_loan = pagibigloan.rate_to_deduct
                                # find pagibig loan contrib by payroll id
                                found_pagibigloan_contrib = Employee_pagibigloan_contrib.objects.filter(
                                    payroll=pay.id).count()
                                if found_pagibigloan_contrib > 0:
                                    # update
                                    found_pagibigloan_contrib_got = Employee_pagibigloan_contrib.objects.get(
                                        payroll=pay.id)
                                    found_pagibigloan_contrib_got.contribution_collected = pagibig_loan
                                    found_pagibigloan_contrib_got.save()
                                    print("saved..")

                                    # check if paid na
                                    total_contributed = Employee_pagibigloan_contrib.objects.filter(
                                        pagibig_loan=pagibigloan.id, employee=pay.employee)
                                    # return HttpResponse(total_contributed)
                                    total_conts = 0
                                    for tots in total_contributed:
                                        total_conts = total_conts + tots.contribution_collected

                                    if total_conts >= pagibigloan.load_amount:
                                        # save status to true
                                        pagibigloan = get_object_or_404(
                                            Employee_pagibig_loan, employee=pay.employee, status=False)
                                        pagibigloan.status = True
                                        pagibigloan.save()

                                else:
                                    # add / insert

                                    # Employee_comloan_contrib
                                    # # save to Employee_comloan_contrib
                                    Employee_pagibigloan_contrib.objects.create(
                                        employee=pay.employee,
                                        cut_off_date=base_payroll.end_date,
                                        pagibig_loan=pagibigloan,
                                        contribution_collected=pagibig_loan,
                                        payroll=pay.id
                                    )

                                    # check if paid na
                                    total_contributed = Employee_pagibigloan_contrib.objects.filter(
                                        pagibig_loan=pagibigloan.id, employee=pay.employee)
                                    # return HttpResponse(total_contributed)
                                    total_conts = 0
                                    for tots in total_contributed:
                                        total_conts = total_conts + tots.contribution_collected

                                    if total_conts >= pagibigloan.load_amount:
                                        # save status to true
                                        pagibigloan = get_object_or_404(
                                            Employee_pagibig_loan, employee=pay.employee, status=False)
                                        pagibigloan.status = True
                                        pagibigloan.save()

                            else:
                                print("sfsdf")

                # check existing vale loan if meron
                found_vale_loans = Employee_vale.objects.filter(
                    employee=pay.employee, status=False)
                print(f'payroll id+ {pay.id}')

                if pay.vale > 0:
                    for valeloan in found_vale_loans:
                        if valeloan.status == False:
                            vale_loan = valeloan.rate_to_deduct
                            # find vale loan contrib by payroll id
                            # found_valeloan_contrib = get_object_or_404(Employee_valeloan_contrib, payroll=pay.id)
                            found_valeloan_contrib = Employee_valeloan_contrib.objects.filter(
                                payroll=pay.id).count()

                            if found_valeloan_contrib > 0:
                                # update
                                found_valeloan_contrib_got = Employee_valeloan_contrib.objects.get(
                                    payroll=pay.id)

                                found_valeloan_contrib_got.contribution_collected = vale_loan
                                found_valeloan_contrib_got.save()
                                print("saved..")
                            else:
                                # add / insert

                                # Employee_comloan_contrib
                                # # save to Employee_comloan_contrib
                                Employee_valeloan_contrib.objects.create(
                                    employee=pay.employee,
                                    cut_off_date=base_payroll.end_date,
                                    vale_loan=valeloan,
                                    contribution_collected=vale_loan,
                                    payroll=pay.id
                                )
                        else:
                            print("sfsdf")

                # check existing canteen loan if meron
                found_canteen_loans = Employee_canteen.objects.filter(
                    employee=pay.employee, status=False)
                print(f'payroll id+ {pay.id}')

                if pay.canteen > 0:
                    for canteen in found_canteen_loans:
                        if canteen.status == False:
                            canteen_loan = canteen.rate_to_deduct
                            # find vale loan contrib by payroll id
                            # found_valeloan_contrib = get_object_or_404(Employee_valeloan_contrib, payroll=pay.id)
                            found_canteen_contrib = Employee_canteen_contrib.objects.filter(
                                payroll=pay.id).count()

                            if found_canteen_contrib > 0:
                                # update
                                found_canteen_contrib_got = Employee_canteen_contrib.objects.get(
                                    payroll=pay.id)

                                found_canteen_contrib_got.contribution_collected = canteen_loan
                                found_canteen_contrib_got.save()
                                print("saved..")
                            else:
                                # add / insert

                                # Employee_comloan_contrib
                                # # save to Employee_comloan_contrib
                                Employee_canteen_contrib.objects.create(
                                    employee=pay.employee,
                                    cut_off_date=base_payroll.end_date,
                                    canteen_loan=canteen,
                                    contribution_collected=canteen_loan,
                                    payroll=pay.id
                                )

                            # check if paid na
                            total_contributed = Employee_canteen_contrib.objects.filter(
                                canteen_loan=canteen.id, employee=pay.employee)
                            # return HttpResponse(total_contributed)
                            total_conts = 0
                            for tots in total_contributed:
                                total_conts = total_conts + tots.contribution_collected

                            if total_conts >= canteen.amount:
                                # save status to true
                                canteen_loan = get_object_or_404(
                                    Employee_canteen, employee=pay.employee, status=False)
                                canteen_loan.status = True
                                canteen_loan.save()
                        else:
                            print("sfsdf")

                # check existing medical loan if meron
                found_medical_loans = Employee_medical.objects.filter(
                    employee=pay.employee, status=False)
                print(f'payroll id+ {pay.id}')

                if pay.medical > 0:
                    for medical in found_medical_loans:
                        if medical.status == False:
                            medical_loan = medical.rate_to_deduct
                            # find vale loan contrib by payroll id
                            # found_valeloan_contrib = get_object_or_404(Employee_valeloan_contrib, payroll=pay.id)
                            found_medical_contrib = Employee_medical_contrib.objects.filter(
                                payroll=pay.id).count()

                            if found_medical_contrib > 0:
                                # update
                                found_medical_contrib_got = Employee_medical_contrib.objects.get(
                                    payroll=pay.id)

                                found_medical_contrib_got.contribution_collected = medical_loan
                                found_medical_contrib_got.save()
                                print("saved..")
                            else:
                                # add / insert

                                # Employee_comloan_contrib
                                # # save to Employee_comloan_contrib
                                Employee_medical_contrib.objects.create(
                                    employee=pay.employee,
                                    cut_off_date=base_payroll.end_date,
                                    medical_loan=medical,
                                    contribution_collected=medical_loan,
                                    payroll=pay.id
                                )
                            # check if paid na
                            total_contributed = Employee_medical_contrib.objects.filter(
                                medical_loan=medical.id, employee=pay.employee)
                            # return HttpResponse(total_contributed)
                            total_conts = 0
                            for tots in total_contributed:
                                total_conts = total_conts + tots.contribution_collected

                            if total_conts >= medical.amount:
                                # save status to true
                                medical_loan = get_object_or_404(
                                    Employee_medical, employee=pay.employee, status=False)
                                medical_loan.status = True
                                medical_loan.save()

                        else:
                            print("sfsdf")

                # check existing gatepass loan if meron
                found_gatepass_loans = Employee_gatepass.objects.filter(
                    employee=pay.employee, status=False)
                print(f'payroll id+ {pay.id}')

                if pay.gatepass > 0:
                    for gatepass in found_gatepass_loans:
                        if gatepass.status == False:
                            gatepass_loan = gatepass.rate_to_deduct
                            # find vale loan contrib by payroll id
                            # found_valeloan_contrib = get_object_or_404(Employee_valeloan_contrib, payroll=pay.id)
                            found_gatepass_contrib = Employee_gatepass_contrib.objects.filter(
                                payroll=pay.id).count()

                            if found_gatepass_contrib > 0:
                                # update
                                found_gatepass_contrib_got = Employee_gatepass_contrib.objects.get(
                                    payroll=pay.id)

                                found_gatepass_contrib_got.contribution_collected = gatepass_loan
                                found_gatepass_contrib_got.save()
                                print("saved..")
                            else:
                                # add / insert

                                # Employee_comloan_contrib
                                # # save to Employee_comloan_contrib
                                Employee_gatepass_contrib.objects.create(
                                    employee=pay.employee,
                                    cut_off_date=base_payroll.end_date,
                                    gatepass_loan=gatepass,
                                    contribution_collected=gatepass_loan,
                                    payroll=pay.id
                                )

                            # check if paid na
                            total_contributed = Employee_gatepass_contrib.objects.filter(
                                gatepass_loan=gatepass.id, employee=pay.employee)
                            # return HttpResponse(total_contributed)
                            total_conts = 0
                            for tots in total_contributed:
                                total_conts = total_conts + tots.contribution_collected

                            if total_conts >= gatepass.amount:
                                # save status to true
                                gatepass_loan = get_object_or_404(
                                    Employee_gatepass, employee=pay.employee, status=False)
                                gatepass_loan.status = True
                                gatepass_loan.save()
                        else:
                            print("sfsdf")


            Logs.objects.create(
                action=f"Payroll for {base_payroll} was sucessfully updated", action_by=request.user, action_date=datetime.now())
            messages.success(request, f'Payroll was successfully saved')
            return redirect('payroll-view', pk=pk)
            # return HttpResponse("sdfds")
        else:
            return HttpResponse(form)
    formset = PayrollFullAddForm(queryset=found_emp_payroll)
    context = {
        'title': 'company',
        'head': f'Payroll View - ({base_payroll})',
        'found_emp_payroll': found_emp_payroll,
        'formset': formset,
        'payroll_view_id': pk,
        'base_payroll': base_payroll,
        'all_employees' : all_employees
    }

    return render(request, 'payrolllist/payroll_list.html', context)
