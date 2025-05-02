from django.shortcuts import render, redirect, get_object_or_404
from .models import Employee, Employee_preferences, Employment_record, Employee_hiring_details, Employee_resume, Employee_sss, Employee_pagibig, Employee_philhealth, Employee_company_loan, Employee_comloan_contrib, Employee_uniform, Employee_medical, Employee_medical_contrib, Employee_canteen, Employee_canteen_contrib, Employee_gatepass, Employee_gatepass_contrib, Employee_vale, Employee_valeloan_contrib, Employee_pagibig_loan, Employee_pagibigloan_contrib, Employee_sss_loan, Employee_sssloan_contrib, Employee_acceptance, Employee_return_to_work, Employee_leave_history, Employee_citizenship, Employee_picture, Employee_requirements, Employee_memo
from django.http import HttpResponse
from .forms import EmployeeAddForm, PreferencesForm, RecordsForm, HiringDetailsForm, ResumeForm, CompanyLoanForm, UniformForm, MedicalForm, CanteenForm, GatepassForm, ValeForm, PagibigLoanForm, SssLoanForm, SearchForm, EmployeeAcceptanceForm, EmployeeReturnToWork, EmployeeLeaveHistory, UpdateRecord, PictureForm, Requirements, MemoForm
from django.contrib import messages
from django.conf import settings
from django.db.models import Q
from hrms.models import Company
from logs.models import Logs
from payrolllist.models import Payroll, Payroll_for_phil_asia
from itertools import chain
import os
import csv
import xlwt
import datetime
from datetime import datetime, timedelta
from django.contrib.auth.decorators import login_required
from django.core.files.storage import FileSystemStorage
from django.conf import settings
from general_settings.models import General_settings


@login_required
def employee_data_movement(request, pk):
    employee = get_object_or_404(Employee, pk=pk)
    logs = Logs.objects.filter(employee=pk)
    context = {
        'head': 'Employee Data Movement - ' + employee.first_name + ' ' + employee.last_name,
        'for_update': 1,
        'pk': pk,
        'employee': employee,
        'logs': logs
    }
    return render(request, 'employee/employee_data_movement.html', context)


@login_required
def memo_delete(request, pk):
    # delete memo
    if request.method == "POST":
        memo = get_object_or_404(Employee_memo, pk=pk)
        employee_pk = memo.employee_id
        path_file = 'employee/media/' + str(memo.memo)
        os.remove(os.path.join(settings.BASE_DIR, path_file))
        Logs.objects.create(employee=memo.employee, action=f"{memo.employee}. employee memo was successfully deleted.",
                                action_by=request.user, action_date=datetime.now())
        memo.delete()

        messages.success(
            request, 'Employee memo was successfully deleted.')
        return redirect('employee-memo', pk=employee_pk)


@login_required
def employee_memo(request, pk):
    employee = get_object_or_404(Employee, pk=pk)
    may_memo = Employee_memo.objects.filter(employee=employee)
    if request.method == 'POST':
        form = MemoForm(request.POST, request.FILES)
        # old_file = employee.memo
        if form.is_valid():
            # return HttpResponse(form)
            # path_file = 'employee/media/memo/' + str(old_file)
            # os.remove(os.path.join(settings.BASE_DIR, path_file))
            f = form.save(commit=False)
            f.employee = employee
            f.save()
            Logs.objects.create(employee=employee, action=f"{employee}. employee memo was successfully uploaded.",
                                action_by=request.user, action_date=datetime.now())
            messages.success(
                request, 'Employee memo was successfully updated.')
            return redirect('employee-memo', pk=pk)
    else:
        form = MemoForm(instance=employee)

    context = {
        'head': 'Update Employee Memo - ' + employee.first_name + ' ' + employee.last_name,
        'form': form,
        'for_update': 1,
        'pk': pk,
        'may_memo': may_memo,
        'employee': employee
    }
    return render(request, 'employee/employee_memo.html', context)


@login_required
def employee_requirements(request, pk):

    employee = get_object_or_404(Employee_requirements, employee_id=pk)
    # return HttpResponse("sdfs")

    if request.method == 'POST':
        form = Requirements(request.POST or None, instance=employee)
        if form.is_valid():
            form.save()
            Logs.objects.create(employee=employee.employee, action=f"{employee.employee}. employee requirements was successfully updated.",
                                action_by=request.user, action_date=datetime.now())
            messages.success(request, "Employee requirements has been updated")
            return redirect("employee-requirements", pk=pk)

    form = Requirements(instance=employee)
    context = {
        'head': 'Update Employee Requirements- ' + employee.employee.first_name + ' ' + employee.employee.last_name,
        'form': form,
        'for_update': 1,
        'pk': pk,
        # 'image_path': image_path,
        'employee': employee.employee
    }
    return render(request, 'employee/requirements.html', context)


@login_required
def employee_upload_picture(request, pk):
    image_path = ""
    employee = get_object_or_404(Employee_picture, employee=pk)
    employee_info = get_object_or_404(Employee, pk=pk)
    if request.method == 'POST':
        file = request.FILES['picture']

        # check if may picture nya..then delete
        if employee.picture:
            os.remove(f"{settings.MEDIA_ROOT}/{employee.picture}")
            print("File Removed!")
            employee.picture = ""
            employee.save()

        fs = FileSystemStorage()
        fs.save(file.name, file)
        employee_picture = get_object_or_404(Employee_picture, employee=pk)
        employee_picture.picture = file.name
        # to_save = Employee_picture(picture=file.name, employee_id = pk)

        #  to_save.save()
        employee_picture.save()

        action = f"Employee {employee_info} successfully uploaded picture."
        Logs.objects.create(employee=employee_info, action=action,
                            action_by=request.user, action_date=datetime.now())

        messages.success(
            request, 'Employee picture was successfully updated.')
        return redirect('employee-upload-picture', pk=pk)
    else:
        form = PictureForm(instance=employee)

        # if employee:
        #   image_path = settings.MEDIA_ROOT+"/"+employee.picture
    context = {
        'head': 'Update Employee Picture- ' + employee.employee.first_name + ' ' + employee.employee.last_name,
        'form': form,
        'for_update': 1,
        'pk': pk,
        # 'image_path': image_path,
        'employee': employee
    }
    return render(request, 'employee/employee_picture.html', context)


@login_required
def employee_endorsement_letter(request, pk):
    employee = Employee.objects.filter(pk=pk).first()
    # return HttpResponse(employee)
    response = HttpResponse(content_type='text/ms-excel')
    file_name = f"{employee} - endorsement letter.xls"
    response['Content-Disposition'] = f'attachment; filename="{file_name}"'
    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('leave absence')
    style_bold = 'align: wrap on, vert centre, horiz center; font: bold on'
    style_square_center = 'align: wrap on, vert center, horiz center; font: bold on, color black; borders: top_color black, bottom_color black, right_color black, left_color black,\
        left thin, right thin, top thin, bottom thin; pattern: pattern solid, fore_color white;'
    style_top_left_right = 'font: height 175; align: wrap on, vert center, horiz center; font: bold on, color black; borders: top_color black, right_color black, left_color black,\
        left thin, right thin, top thin; pattern: pattern solid, fore_color white;'
    style_left = 'font: height 150; align: wrap on, vert center, horiz center; font: bold on, color black; borders: left_color black,\
        left thin; pattern: pattern solid, fore_color white;'
    style_right = 'font: height 150; align: wrap on, vert center, horiz center; font: color black; borders: right_color black,\
        right thin; pattern: pattern solid, fore_color white;'
    style_left_right = 'font: height 175; align: wrap on, vert center, horiz center; font: bold on, color black; borders: left_color black, right_color black,\
        left thin, right thin; pattern: pattern solid, fore_color white;'
    style_bottom_right = 'font: height 175; align: wrap on, vert center, horiz center; font: bold on, color black; borders: bottom_color black, right_color black,\
        right thin, bottom thin; pattern: pattern solid, fore_color white;'
    style_bottom = 'font: height 175; align: wrap on, vert center, horiz center; font: bold on, color black; borders: bottom_color black,\
        bottom thin; pattern: pattern solid, fore_color white;'

    gen_settings = General_settings.objects.get(id=1)
    template_name = gen_settings.template_name
    company_address = gen_settings.company_address
    company_contacts = gen_settings.company_contacts


    ws.write_merge(0, 1, 0, 10, template_name,
                   xlwt.Style.easyxf(style_bold))
    style_normal = 'font: height 175; align: wrap on, vert centre, horiz center;'
    ws.write_merge(2, 2, 0, 10, company_address,
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(3, 3, 0, 10, company_contacts,
                   xlwt.Style.easyxf(style_normal))

    ws.write_merge(5, 5, 0, 10, 'ENDORSEMENT LETTER',
                   xlwt.Style.easyxf(style_square_center))

    ws.write_merge(7, 7, 0, 0, 'DATE:',
                   xlwt.Style.easyxf(style_normal))

    ws.write_merge(9, 9, 0, 0, 'ATTENTION',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(9, 9, 1, 5, '',
                   xlwt.Style.easyxf(style_bottom))
    ws.write_merge(11, 11, 0, 0, 'SUBJECT',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(11, 11, 1, 5, '',
                   xlwt.Style.easyxf(style_bottom))

    ws.write_merge(14, 14, 0, 0, 'Dear',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(14, 14, 1, 5, '',
                   xlwt.Style.easyxf(style_bottom))

    ws.write_merge(16, 16, 0, 1, 'Please accomodate ',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(16, 16, 2, 4, f'{employee}',
                   xlwt.Style.easyxf(style_bottom))
    ws.write_merge(16, 16, 5, 9, ' an (applicant, trainee, apprentice) subject for',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(17, 17, 0, 2, ' (INTERVIEW) (EXAM).',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(19, 19, 0, 6, 'Give us feedback whatever the result of his / her initial assessment with them.',
                   xlwt.Style.easyxf(style_normal))

    ws.write_merge(22, 22, 0, 1, 'Thank you very much.',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(24, 24, 0, 1, 'Sinceryly yours,',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(26, 26, 0, 1, '',
                   xlwt.Style.easyxf(style_bottom))
    Logs.objects.create(employee=employee, action=f"{employee}. Downloaded endorsement letter.",
                        action_by=request.user, action_date=datetime.now())
    wb.save(response)
    return response


@login_required
def employee_leave_absence(request, pk):
    employee = get_object_or_404(Employee, pk=pk)
    may_leave = Employee_leave_history.objects.filter(employee=pk)
    if request.method == 'POST':
        form = EmployeeLeaveHistory(request.POST)
        if form.is_valid():
            la_save = form.save(commit=False)
            la_save.employee_id = pk
            la_save.save()

            Logs.objects.create(employee=employee, action=f"{employee}. Added leave of absence record (start_date={form.cleaned_data['start_date']}, end_date={form.cleaned_data['end_date']}, no_of_days={form.cleaned_data['no_of_days']})",
                                action_by=request.user, action_date=datetime.now())
            messages.success(
                request, 'Employmee leave details was successfully saved.')
            return redirect('employee-leave-absence', pk=pk)
    else:
        form = EmployeeLeaveHistory(instance=employee)
    context = {
        'head': 'Update Employee - ' + employee.first_name + ' ' + employee.last_name + " Leave of Absence",
        'form': form,
        'for_update': 1,
        'pk': pk,
        'may_leave': may_leave,
        'employee': employee,
    }
    return render(request, 'employee/employee_leave_absence.html', context)


@login_required
def employee_leave_download(request, emp_pk, leave_pk):
    employee = Employee.objects.filter(pk=emp_pk).first()
    # return HttpResponse(employee)
    leave = Employee_leave_history.objects.filter(pk=leave_pk).first()
    response = HttpResponse(content_type='text/ms-excel')
    file_name = f"{employee} - Leave Absence.xls"
    response['Content-Disposition'] = f'attachment; filename="{file_name}"'
    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('leave absence')
    style_bold = 'align: wrap on, vert centre, horiz center; font: bold on'
    style_square_center = 'align: wrap on, vert center, horiz center; font: bold on, color black; borders: top_color black, bottom_color black, right_color black, left_color black,\
        left thin, right thin, top thin, bottom thin; pattern: pattern solid, fore_color white;'
    style_top_left_right = 'font: height 175; align: wrap on, vert center, horiz center; font: bold on, color black; borders: top_color black, right_color black, left_color black,\
        left thin, right thin, top thin; pattern: pattern solid, fore_color white;'
    style_left = 'font: height 150; align: wrap on, vert center, horiz center; font: bold on, color black; borders: left_color black,\
        left thin; pattern: pattern solid, fore_color white;'
    style_right = 'font: height 150; align: wrap on, vert center, horiz center; font: color black; borders: right_color black,\
        right thin; pattern: pattern solid, fore_color white;'
    style_left_right = 'font: height 175; align: wrap on, vert center, horiz center; font: bold on, color black; borders: left_color black, right_color black,\
        left thin, right thin; pattern: pattern solid, fore_color white;'
    style_bottom_right = 'font: height 175; align: wrap on, vert center, horiz center; font: bold on, color black; borders: bottom_color black, right_color black,\
        right thin, bottom thin; pattern: pattern solid, fore_color white;'
    style_bottom = 'font: height 175; align: wrap on, vert center, horiz center; font: bold on, color black; borders: bottom_color black,\
        bottom thin; pattern: pattern solid, fore_color white;'

    gen_settings = General_settings.objects.get(id=1)
    template_name = gen_settings.template_name
    company_address = gen_settings.company_address
    company_contacts = gen_settings.company_contacts

    ws.write_merge(0, 1, 0, 10, template_name,
                   xlwt.Style.easyxf(style_bold))
    style_normal = 'font: height 175; align: wrap on, vert centre, horiz center;'
    ws.write_merge(2, 2, 0, 10, company_address,
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(3, 3, 0, 10, company_contacts,
                   xlwt.Style.easyxf(style_normal))

    ws.write_merge(5, 5, 0, 10, 'REQUEST FOR LEAVE ABSENCE',
                   xlwt.Style.easyxf(style_square_center))

    ws.write_merge(6, 6, 0, 10, '',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(7, 7, 0, 0, 'Name:',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(7, 7, 1, 4, f'{employee}',
                   xlwt.Style.easyxf(style_bottom))

    ws.write_merge(7, 7, 6, 7, 'Date Filed:',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(7, 7, 8, 10, f'',
                   xlwt.Style.easyxf(style_bottom))

    ws.write_merge(8, 8, 0, 10, '',
                   xlwt.Style.easyxf(style_normal))

    ws.write_merge(9, 9, 0, 1, 'Type of Leave:',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(9, 9, 2, 2, '',
                   xlwt.Style.easyxf(style_bottom))
    ws.write_merge(9, 9, 3, 3, f'Vacation',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(9, 9, 5, 5, '',
                   xlwt.Style.easyxf(style_bottom))
    ws.write_merge(9, 9, 6, 6, f'Sick',
                   xlwt.Style.easyxf(style_normal))

    ws.write_merge(9, 9, 8, 8, '',
                   xlwt.Style.easyxf(style_bottom))
    ws.write_merge(9, 9, 9, 9, f'Others',
                   xlwt.Style.easyxf(style_normal))

    ws.write_merge(11, 11, 0, 2, f'Date(s) of Leave: From',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(11, 11, 3, 4, f'{leave.start_date}',
                   xlwt.Style.easyxf(style_bottom))
    ws.write_merge(11, 11, 5, 5, f'to',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(11, 11, 6, 7, f'{leave.end_date}',
                   xlwt.Style.easyxf(style_bottom))
    ws.write_merge(11, 11, 8, 9, f'No. of Day(s):',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(11, 11, 10, 10, f'{leave.no_of_days}',
                   xlwt.Style.easyxf(style_bottom))

    ws.write_merge(13, 13, 0, 0, f'Reason(s):',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(13, 13, 1, 10, f'',
                   xlwt.Style.easyxf(style_bottom))
    ws.write_merge(14, 14, 0, 10, f'',
                   xlwt.Style.easyxf(style_bottom))
    ws.write_merge(16, 16, 8, 10, f'',
                   xlwt.Style.easyxf(style_bottom))
    ws.write_merge(17, 17, 8, 10, f'Employee Signature',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(18, 18, 0, 10, f'',
                   xlwt.Style.easyxf(style_bottom))
    ws.write_merge(20, 20, 0, 1, f'Action Taken:',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(20, 20, 2, 2, f'',
                   xlwt.Style.easyxf(style_bottom))
    ws.write_merge(20, 20, 3, 3, f'Approved',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(20, 20, 5, 5, f'',
                   xlwt.Style.easyxf(style_bottom))
    ws.write_merge(20, 20, 6, 6, f'Disapproved',
                   xlwt.Style.easyxf(style_normal))

    ws.write_merge(21, 21, 0, 0, f'Comments:',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(21, 21, 1, 10, f'',
                   xlwt.Style.easyxf(style_bottom))
    ws.write_merge(22, 22, 0, 10, f'',
                   xlwt.Style.easyxf(style_bottom))

    ws.write_merge(24, 24, 0, 2, f'',
                   xlwt.Style.easyxf(style_bottom))
    ws.write_merge(25, 25, 0, 2, f'Immediate Superior',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(24, 24, 4, 6, f'',
                   xlwt.Style.easyxf(style_bottom))
    ws.write_merge(25, 25, 4, 6, f'Department Head',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(24, 24, 8, 10, f'',
                   xlwt.Style.easyxf(style_bottom))
    ws.write_merge(25, 25, 8, 10, f'HRD',
                   xlwt.Style.easyxf(style_normal))


    Logs.objects.create(employee=employee, action=f"{employee}. Downloaded leave of absence form.",
                        action_by=request.user, action_date=datetime.now())
    wb.save(response)
    return response


@login_required
def employee_clearance_form(request, pk):
    employee = Employee.objects.filter(pk=pk).first()
    response = HttpResponse(content_type='text/ms-excel')
    file_name = f"{employee} - Clearance Form.xls"
    response['Content-Disposition'] = f'attachment; filename="{file_name}"'
    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('return to work')
    style_bold = 'align: wrap on, vert centre, horiz center; font: bold on'
    style_square_center = 'align: wrap on, vert center, horiz center; font: bold on, color black; borders: top_color black, bottom_color black, right_color black, left_color black,\
        left thin, right thin, top thin, bottom thin; pattern: pattern solid, fore_color white;'
    style_top_left_right = 'font: height 175; align: wrap on, vert center, horiz center; font: bold on, color black; borders: top_color black, right_color black, left_color black,\
        left thin, right thin, top thin; pattern: pattern solid, fore_color white;'
    style_left = 'font: height 150; align: wrap on, vert center, horiz center; font: bold on, color black; borders: left_color black,\
        left thin; pattern: pattern solid, fore_color white;'
    style_right = 'font: height 150; align: wrap on, vert center, horiz center; font: color black; borders: right_color black,\
        right thin; pattern: pattern solid, fore_color white;'
    style_left_right = 'font: height 175; align: wrap on, vert center, horiz center; font: bold on, color black; borders: left_color black, right_color black,\
        left thin, right thin; pattern: pattern solid, fore_color white;'
    style_bottom_right = 'font: height 175; align: wrap on, vert center, horiz center; font: bold on, color black; borders: bottom_color black, right_color black,\
        right thin, bottom thin; pattern: pattern solid, fore_color white;'
    style_bottom = 'font: height 175; align: wrap on, vert center, horiz center; font: bold on, color black; borders: bottom_color black,\
        bottom thin; pattern: pattern solid, fore_color white;'

    gen_settings = General_settings.objects.get(id=1)
    template_name = gen_settings.template_name
    company_name = gen_settings.main_company
    company_address = gen_settings.company_address
    company_contacts = gen_settings.company_contacts

    ws.write_merge(0, 1, 0, 10, template_name,
                   xlwt.Style.easyxf(style_bold))
    style_normal = 'font: height 175; align: wrap on, vert centre, horiz center;'
    ws.write_merge(2, 2, 0, 10, company_address,
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(3, 3, 0, 10, company_contacts,
                   xlwt.Style.easyxf(style_normal))

    ws.write_merge(5, 5, 0, 10, 'Clearance Form',
                   xlwt.Style.easyxf(style_square_center))

    ws.write_merge(7, 7, 0, 4, 'THIS PORTION IS TO BE FILLED UP BY EMPLOYEE',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(8, 8, 0, 4, '',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(9, 9, 0, 0, 'Date:',
                   xlwt.Style.easyxf(style_left))
    ws.write_merge(9, 9, 1, 4, f'{datetime.now().strftime("%Y-%m-%d")}',
                   xlwt.Style.easyxf(style_bottom))
    ws.write_merge(10, 10, 0, 1, 'CLIENT COMPANY:',
                   xlwt.Style.easyxf(style_left))
    ws.write_merge(10, 10, 2, 4, f'{employee.company}',
                   xlwt.Style.easyxf(style_bottom))
    ws.write_merge(11, 11, 0, 0, f'NAME:',
                   xlwt.Style.easyxf(style_left))
    ws.write_merge(11, 11, 1, 4, f'{employee}',
                   xlwt.Style.easyxf(style_bottom))
    ws.write_merge(12, 12, 0, 0, f'SIGNATURE:',
                   xlwt.Style.easyxf(style_left))
    ws.write_merge(12, 12, 1, 4, f'',
                   xlwt.Style.easyxf(style_bottom))
    ws.write_merge(13, 13, 0, 0, f'CONTACT:',
                   xlwt.Style.easyxf(style_left))
    ws.write_merge(13, 13, 1, 4, f'{employee.phone}',
                   xlwt.Style.easyxf(style_bottom))
    ws.write_merge(14, 14, 0, 4, f'',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(15, 15, 0, 2, f'This certifies that Mr./Ms.',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(15, 15, 3, 4, f'{employee}',
                   xlwt.Style.easyxf(style_bottom))
    ws.write_merge(16, 16, 0, 4, f'is free from any liability, monetary or otherwise with',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(17, 17, 0, 4, f'{ template_name.upper() } as of the date and time',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(18, 18, 0, 4, f'opposite our affixed signature and as of the close of working',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(19, 19, 0, 4, f'hours on (last day).',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(24, 24, 0, 10, f'',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(25, 25, 0, 5, f'RESIGNATION LETTER',
                   xlwt.Style.easyxf(style_top_left_right))
    ws.write_merge(26, 26, 0, 5, f'',
                   xlwt.Style.easyxf(style_left_right))
    ws.write_merge(27, 27, 0, 5, f'',
                   xlwt.Style.easyxf(style_square_center))
    ws.write_merge(28, 28, 0, 5, f'',
                   xlwt.Style.easyxf(style_square_center))
    ws.write_merge(29, 29, 0, 5, f'',
                   xlwt.Style.easyxf(style_square_center))
    ws.write_merge(30, 30, 0, 5, f'',
                   xlwt.Style.easyxf(style_square_center))
    ws.write_merge(31, 31, 0, 5, f'',
                   xlwt.Style.easyxf(style_square_center))
    ws.write_merge(32, 32, 0, 5, f'',
                   xlwt.Style.easyxf(style_square_center))
    ws.write_merge(33, 33, 0, 5, f'',
                   xlwt.Style.easyxf(style_square_center))
    ws.write_merge(34, 34, 0, 5, f'',
                   xlwt.Style.easyxf(style_square_center))
    ws.write_merge(25, 26, 6, 10, f'THIS PORTION TO BE FILLED UP BY { company_name.upper() }',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(27, 27, 6, 7, f'RELEASED BY:',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(27, 27, 8, 10, f'',
                   xlwt.Style.easyxf(style_bottom))
    ws.write_merge(28, 28, 6, 7, f'DATE:',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(28, 28, 8, 10, f'',
                   xlwt.Style.easyxf(style_bottom))
    ws.write_merge(29, 29, 6, 10, f'',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(30, 30, 6, 7, f'RECEIVED BY:',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(30, 30, 8, 10, f'',
                   xlwt.Style.easyxf(style_bottom))
    ws.write_merge(31, 31, 6, 7, f'DATE:',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(31, 31, 8, 10, f'',
                   xlwt.Style.easyxf(style_bottom))

    # Right square
    ws.write_merge(7, 7, 6, 10, 'THIS PORTION IS TO BE FILLED UP BY COMPANY',
                   xlwt.Style.easyxf(style_top_left_right))
    ws.write_merge(8, 8, 6, 10, '',
                   xlwt.Style.easyxf(style_left_right))
    ws.write_merge(9, 9, 6, 7, 'RECEIVED BY:',
                   xlwt.Style.easyxf(style_left))
    ws.write_merge(9, 9, 8, 10, '',
                   xlwt.Style.easyxf(style_bottom_right))
    ws.write_merge(10, 10, 6, 7, 'DATE:',
                   xlwt.Style.easyxf(style_left))
    ws.write_merge(10, 10, 8, 10, '',
                   xlwt.Style.easyxf(style_bottom_right))
    ws.write_merge(11, 11, 6, 7, 'EVALUATED BY:',
                   xlwt.Style.easyxf(style_left))
    ws.write_merge(11, 11, 8, 10, '',
                   xlwt.Style.easyxf(style_bottom_right))
    ws.write_merge(12, 12, 6, 7, 'DATE:',
                   xlwt.Style.easyxf(style_left))
    ws.write_merge(12, 12, 8, 10, '',
                   xlwt.Style.easyxf(style_bottom_right))
    ws.write_merge(13, 13, 6, 10, '',
                   xlwt.Style.easyxf(style_left_right))
    ws.write_merge(14, 14, 6, 10, '',
                   xlwt.Style.easyxf(style_left_right))
    ws.write_merge(15, 15, 6, 6, 'Remarks:',
                   xlwt.Style.easyxf(style_left))
    ws.write_merge(15, 15, 7, 10, '',
                   xlwt.Style.easyxf(style_right))
    ws.write_merge(16, 16, 6, 6, '',
                   xlwt.Style.easyxf(style_square_center))
    ws.write_merge(16, 16, 7, 7, 'AWOL',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(16, 16, 8, 8, '',
                   xlwt.Style.easyxf(style_square_center))
    ws.write_merge(16, 16, 9, 10, 'FINISHED CONTRACT',
                   xlwt.Style.easyxf(style_right))

    ws.write_merge(17, 17, 6, 10, '',
                   xlwt.Style.easyxf(style_left_right))

    ws.write_merge(18, 18, 6, 6, '',
                   xlwt.Style.easyxf(style_square_center))
    ws.write_merge(18, 18, 7, 7, 'TERMINATED',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(18, 18, 8, 8, '',
                   xlwt.Style.easyxf(style_square_center))
    ws.write_merge(18, 18, 9, 10, 'RESIGNED',
                   xlwt.Style.easyxf(style_right))

    ws.write_merge(19, 19, 6, 10, '',
                   xlwt.Style.easyxf(style_left_right))
    ws.write_merge(20, 20, 6, 6, 'COMMENTS',
                   xlwt.Style.easyxf(style_left))
    ws.write_merge(20, 20, 7, 10, '',
                   xlwt.Style.easyxf(style_right))
    ws.write_merge(21, 21, 6, 10, '',
                   xlwt.Style.easyxf(style_square_center))
    ws.write_merge(22, 22, 6, 10, '',
                   xlwt.Style.easyxf(style_square_center))
    ws.write_merge(23, 23, 6, 10, '',
                   xlwt.Style.easyxf(style_square_center))

    ws.write_merge(35, 35, 0, 10, '',
                   xlwt.Style.easyxf(style_bottom))
    ws.write_merge(36, 36, 0, 10, '',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(37, 37, 0, 4, 'THIS PORTION TO BE FILLED UP BY '+company_name.upper(),
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(38, 38, 0, 10, '',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(39, 39, 0, 10, 'CLAIMING STUB',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(40, 40, 0, 10, '',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(41, 41, 0, 0, 'NAME:',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(41, 41, 1, 4, F'{employee}',
                   xlwt.Style.easyxf(style_bottom))
    ws.write_merge(41, 41, 6, 7, F'ID SURRENDERED:',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(41, 41, 8, 8, F'',
                   xlwt.Style.easyxf(style_square_center))
    ws.write_merge(42, 42, 0, 1, 'DATE RELEASED:',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(42, 42, 2, 4, F'',
                   xlwt.Style.easyxf(style_bottom))
    ws.write_merge(42, 42, 6, 6, F'VOUCHER:',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(42, 42, 7, 7, F'',
                   xlwt.Style.easyxf(style_square_center))
    ws.write_merge(42, 42, 8, 9, F'PAYSLIP SURRENDERED',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(42, 42, 10, 10, F'',
                   xlwt.Style.easyxf(style_square_center))
    ws.write_merge(43, 43, 10, 10, F'',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(44, 44, 0, 10, F'NOTE: BRING THIS COUPON & VALID ID TO CLAIM YOUR CASH BOND. STRICTLY NO REPRESENTATIVE.',
                   xlwt.Style.easyxf(style_normal))
    Logs.objects.create(employee=employee, action=f"{employee}. Downloaded clearance form.",
                        action_by=request.user, action_date=datetime.now())
    wb.save(response)
    return response


@login_required
def employee_return_download(request, pk):
    employment_record = Employment_record.objects.filter(employee=pk).first()
    employee_return_to_work = get_object_or_404(
        Employee_return_to_work, employee=pk)
    response = HttpResponse(content_type='text/ms-excel')
    file_name = f"{employee_return_to_work.employee} - Return to work.xls"
    response['Content-Disposition'] = f'attachment; filename="{file_name}"'
    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('return to work')
    style = 'align: wrap on, vert centre, horiz center; font: bold on'
#    ws.row(0).write(0, value, xlwt.Style.easyxf(style))

    gen_settings = General_settings.objects.get(id=1)
    template_name = gen_settings.template_name
    company_name = gen_settings.main_company
    company_address = gen_settings.company_address
    company_contacts = gen_settings.company_contacts

    ws.write_merge(0, 1, 0, 9, template_name,
                   xlwt.Style.easyxf(style))
    style = 'align: wrap on, vert centre, horiz center;'
    ws.write_merge(2, 2, 0, 9, company_address,
                   xlwt.Style.easyxf(style))
    ws.write_merge(3, 3, 0, 9, company_contacts,
                   xlwt.Style.easyxf(style))
    style = 'align: wrap on, vert centre, horiz center; font: bold on'
    ws.write_merge(5, 5, 7, 7, 'Date: ',
                   xlwt.Style.easyxf(style))
    ws.write_merge(5, 5, 8, 9, f'{datetime.now().strftime("%Y-%m-%d")}',
                   xlwt.Style.easyxf(style))

    style_ordinary = 'align: wrap on; font: bold on'
    ws.write_merge(6, 6, 0, 0, 'Name: ',
                   xlwt.Style.easyxf(style_ordinary))
    ws.write_merge(6, 6, 1, 4, f'{employee_return_to_work.employee}',
                   xlwt.Style.easyxf(style_ordinary))
    ws.write_merge(7, 7, 0, 0, 'Company: ',
                   xlwt.Style.easyxf(style_ordinary))
    ws.write_merge(7, 7, 1, 4, f'{employee_return_to_work.employee.company}',
                   xlwt.Style.easyxf(style_ordinary))
    ws.write_merge(8, 8, 0, 0, 'Postion: ',
                   xlwt.Style.easyxf(style_ordinary))
    ws.write_merge(8, 8, 1, 4, f'{employment_record.position}',
                   xlwt.Style.easyxf(style_ordinary))

    ws.write_merge(10, 10, 0, 12, f'Please let this letter serve as notification that I will return to work on {employee_return_to_work.return_on}. I began my ',
                   xlwt.Style.easyxf(style_ordinary))
    ws.write_merge(11, 11, 0, 12, f'leave/ absence on {employee_return_to_work.absence_date_start} to {employee_return_to_work.absence_date_end}.',
                   xlwt.Style.easyxf(style_ordinary))
    ws.write_merge(12, 12, 0, 12, f'REASON: {employee_return_to_work.reason}',
                   xlwt.Style.easyxf(style_ordinary))
    ws.write_merge(14, 14, 0, 2, f'Approved by:',
                   xlwt.Style.easyxf(style_ordinary))
    ws.write_merge(15, 15, 0, 2, f'__________________',
                   xlwt.Style.easyxf(style_ordinary))
    ws.write_merge(15, 15, 8, 9, f'__________________',
                   xlwt.Style.easyxf(style_ordinary))

    ws.write_merge(16, 16, 0, 2, f'Supervisor Head',
                   xlwt.Style.easyxf(style_ordinary))
    ws.write_merge(16, 16, 8, 9, f'HR Department',
                   xlwt.Style.easyxf(style_ordinary))

    ws.write_merge(17, 17, 0, 2, f'__________________',
                   xlwt.Style.easyxf(style_ordinary))
    ws.write_merge(17, 17, 7, 9, f'_____________________________',
                   xlwt.Style.easyxf(style_ordinary))

    ws.write_merge(18, 18, 0, 2, f'{ company_name.title() }',
                   xlwt.Style.easyxf(style_ordinary))
    ws.write_merge(18, 18, 7, 9, f'Employee\'s Signature',
                   xlwt.Style.easyxf(style_ordinary))

    Logs.objects.create(employee=employee_return_to_work.employee, action=f"{employee_return_to_work.employee}. Downloaded return to work form.",
                        action_by=request.user, action_date=datetime.now())
    wb.save(response)
    return response


@login_required
def employee_return_to_work(request, pk):
    employee = Employee_return_to_work.objects.filter(employee=pk).first()
    employee_hiring_details = get_object_or_404(
        Employee_hiring_details, employee=pk)

    if request.method == 'POST':
        employee_return_to_work = EmployeeReturnToWork(
            request.POST or None, instance=employee)
        if employee_return_to_work.is_valid():
            ea = employee_return_to_work.save(commit=False)
            ea.employee_id = pk
            ea.save()

            action = f"{employee_hiring_details.employee}. return to work letter was successfully updated. (return_on - {employee_return_to_work.cleaned_data['return_on']}, absence_date_start - {employee_return_to_work.cleaned_data['absence_date_start']}, absence_date_end - {employee_return_to_work.cleaned_data['absence_date_end']}, reason - {employee_return_to_work.cleaned_data['reason']})"
            Logs.objects.create(employee=employee_hiring_details.employee, action=action,
                                action_by=request.user, action_date=datetime.now())
            messages.success(
                request, f'Employee return to work form details has been successfully saved.')
            return redirect('/employees/employee-return-to-work/' + str(pk))
    else:
        # return HttpResponse("sfsd")
        employee_return_to_work = EmployeeReturnToWork(instance=employee)

    context = {
        'head': 'Add Return To Work Form Details',
        'form': employee_return_to_work,
        'employee_id': pk,
        'pk': pk,
        'for_update': 1,
        'employee': employee_hiring_details.employee
    }
    return render(request, 'employee/return_to_work.html', context)


@login_required
def employee_acceptance_download(request, pk):
    employee = get_object_or_404(Employee_acceptance, employee=pk)
    employee_hiring_details = get_object_or_404(
        Employee_hiring_details, employee=pk)
    response = HttpResponse(content_type='text/ms-excel')
    file_name = f"{employee.employee} - Acceptance.xls"
    response['Content-Disposition'] = f'attachment; filename="{file_name}"'
    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('acceptance')
    style = 'align: wrap on, vert centre, horiz center; font: bold on'
#    ws.row(0).write(0, value, xlwt.Style.easyxf(style))

    gen_settings = General_settings.objects.get(id=1)
    template_name = gen_settings.template_name
    company_name = gen_settings.main_company
    company_address = gen_settings.company_address
    company_contacts = gen_settings.company_contacts

    ws.write_merge(0, 1, 0, 9, template_name.title(),
                   xlwt.Style.easyxf(style))
    style = 'align: wrap on, vert centre, horiz center;'
    ws.write_merge(2, 2, 0, 9, company_address,
                   xlwt.Style.easyxf(style))
    ws.write_merge(3, 3, 0, 9, company_contacts,
                   xlwt.Style.easyxf(style))
    style = 'align: wrap on, vert centre, horiz center; font: bold on'
    ws.write_merge(5, 5, 0, 0, 'Petsa.',
                   xlwt.Style.easyxf(style))
    ws.write_merge(5, 5, 1, 4, f'{datetime.now().strftime("%Y-%m-%d")}',
                   xlwt.Style.easyxf(style))
    ws.write_merge(6, 6, 0, 0, 'G/Bb/Gng.',
                   xlwt.Style.easyxf(style))
    ws.write_merge(6, 6, 1, 5, '',
                   xlwt.Style.easyxf(style))
    ws.write_merge(8, 8, 0, 8, 'Ito ay katunayan na ang pangalan mo ay ibinilang na sa listahan ng aming mga kawani.',
                   xlwt.Style.easyxf(style))
    style = 'align: wrap on; font: bold on'
    ws.write_merge(10, 10, 0, 9, 'Kaugnay nito, ikinagagalak namin ipa-alam sa iyo na isa ka sa mga natanggap namin at',
                   xlwt.Style.easyxf(style))
    ws.write_merge(11, 11, 0, 9, 'nagpapasailalim ka sa mga sumusunod na kondisyon:',
                   xlwt.Style.easyxf(style))
    ws.write_merge(13, 13, 1, 3, '1. Katayuan bilang kawani',
                   xlwt.Style.easyxf(style))
    ws.write_merge(14, 14, 1, 3, '(Employment Status)',
                   xlwt.Style.easyxf(style))

    style_square = 'font: bold on, color black; borders: top_color black, bottom_color black, right_color black, left_color black,\
                              left thin, right thin, top thin, bottom thin; pattern: pattern solid, fore_color white;'
    ws.write_merge(14, 14, 5, 5, '',
                   xlwt.Style.easyxf(style_square))
    style_ordinary = 'align: wrap on; font: bold on'
    ws.write_merge(14, 14, 6, 9, ' Casual / Probationary',
                   xlwt.Style.easyxf(style_ordinary))
    ws.write_merge(16, 16, 5, 5, '',
                   xlwt.Style.easyxf(style_square))
    ws.write_merge(16, 16, 6, 9, ' Seasonal',
                   xlwt.Style.easyxf(style_ordinary))
    ws.write_merge(18, 18, 5, 5, '',
                   xlwt.Style.easyxf(style_square))
    ws.write_merge(18, 18, 6, 9, ' Regular',
                   xlwt.Style.easyxf(style_ordinary))

    ws.write_merge(20, 20, 1, 3, '2. Panahon',
                   xlwt.Style.easyxf(style_ordinary))
    ws.write_merge(20, 20, 5, 5, 'Mula',
                   xlwt.Style.easyxf(style_ordinary))
    ws.write_merge(20, 20, 6, 9, f'{employee.employee.date_hired}',
                   xlwt.Style.easyxf('borders: top_color white, bottom_color black, right_color white, left_color white,\
                              left thin, right thin, top thin, bottom thin;'))
    ws.write_merge(22, 22, 5, 5, 'Hanggang',
                   xlwt.Style.easyxf(style_ordinary))
    ws.write_merge(22, 22, 6, 9, f'{employee.employee.contract_expiration}',
                   xlwt.Style.easyxf('borders: top_color white, bottom_color black, right_color white, left_color white,\
                              left thin, right thin, top thin, bottom thin;'))
    ws.write_merge(23, 23, 5, 9, 'O lagpas pa ayon sa pangangailangan',
                   xlwt.Style.easyxf(style_ordinary))
    ws.write_merge(24, 24, 5, 9, 'aking kliyente',
                   xlwt.Style.easyxf(style_ordinary))

    ws.write_merge(26, 26, 1, 3, 'Puwesto (Position)',
                   xlwt.Style.easyxf(style_ordinary))
    ws.write_merge(26, 26, 5, 7, f'{employee_hiring_details.position}',
                   xlwt.Style.easyxf('borders: top_color white, bottom_color black, right_color white, left_color white,\
                              left thin, right thin, top thin, bottom thin;'))
    ws.write_merge(28, 28, 1, 3, 'Suweldo (Salary',
                   xlwt.Style.easyxf(style_ordinary))
    ws.write_merge(28, 28, 5, 6, f'{employee.salary_per_day}',
                   xlwt.Style.easyxf(style_ordinary))
    ws.write_merge(28, 28, 7, 8, f'bawat araw',
                   xlwt.Style.easyxf('borders: top_color white, bottom_color black, right_color white, left_color white,\
                              left thin, right thin, top thin, bottom thin;'))
    ws.write_merge(30, 30, 5, 6, 'at babayaran',
                   xlwt.Style.easyxf(style_ordinary))
    ws.write_merge(30, 30, 7, 9, f'{employee.salary_per_month} bawat buwan',
                   xlwt.Style.easyxf('borders: top_color white, bottom_color black, right_color white, left_color white,\
                              left thin, right thin, top thin, bottom thin;'))

    ws.write_merge(32, 32, 0, 9, 'Gayundin, nananatili sa amin ang karapatan na ikaw ay alisin sa trabaho ng walang pasubali at',
                   xlwt.Style.easyxf(style_ordinary))
    ws.write_merge(33, 33, 0, 9, 'walang pananagutan sa iyo hinggil sa iyong pagkakaalis sa anumang panahon, kapag nangyari o',
                   xlwt.Style.easyxf(style_ordinary))
    ws.write_merge(34, 34, 0, 9, 'naganap ang alinman sa mga sumusunod:',
                   xlwt.Style.easyxf(style_ordinary))

    ws.write_merge(36, 36, 1, 9, '1. Kapag hindi na kailangan ang iyong serbisyo dahil sa kakulangan ng materyales sa',
                   xlwt.Style.easyxf(style_ordinary))
    ws.write_merge(37, 37, 1, 9, 'paggawa o kaukulang namimili o customer.',
                   xlwt.Style.easyxf(style_ordinary))
    ws.write_merge(38, 38, 1, 9, '2. Kapag hindi kasiya-siya ang iyong serbisyo o kulang ang iyong kwalipikasyon sa iyong',
                   xlwt.Style.easyxf(style_ordinary))
    ws.write_merge(39, 39, 1, 9, 'trabaho.',
                   xlwt.Style.easyxf(style_ordinary))
    ws.write_merge(40, 40, 1, 9, '3. Kapag lumabag ka sa alinman sa mga alintuntunin ng aming kumpanya.',
                   xlwt.Style.easyxf(style_ordinary))

    ws.write_merge(43, 43, 6, 9, 'Nabasa ko at naintindihan at ako ay',
                   xlwt.Style.easyxf(style_ordinary))
    ws.write_merge(44, 44, 6, 9, 'sumasangayon sa mga nabanggit na',
                   xlwt.Style.easyxf(style_ordinary))
    ws.write_merge(45, 45, 6, 9, 'kondisyon ukol sa aking pinsukan',
                   xlwt.Style.easyxf(style_ordinary))

    ws.write_merge(47, 47, 6, 9, '',
                   xlwt.Style.easyxf('borders: top_color white, bottom_color black, right_color white, left_color white,\
                              left thin, right thin, top thin, bottom thin;'))
    ws.write_merge(48, 48, 6, 9, 'Signature (above printed name)',
                   xlwt.Style.easyxf(style_ordinary))

    Logs.objects.create(employee=employee.employee, action=f"{employee.employee}. Downloaded acceptance form.",
                        action_by=request.user, action_date=datetime.now())

    wb.save(response)
    return response


@login_required
def employee_acceptance(request, pk):
    employee = Employee_acceptance.objects.filter(employee=pk).first()

    if request.method == 'POST':
        employee_acceptance = EmployeeAcceptanceForm(
            request.POST or None, instance=employee)
        if employee_acceptance.is_valid():
            ea = employee_acceptance.save(commit=False)
            ea.employee_id = pk
            ea.save()

            action = f"{employee.employee}. acceptance letter was successfully updated. (employment_status - {employee_acceptance.cleaned_data['employment_status']}, start_date - {employee_acceptance.cleaned_data['start_date']}, end_date - {employee_acceptance.cleaned_data['end_date']}, position - {employee_acceptance.cleaned_data['position']}, salary_per_day - {employee_acceptance.cleaned_data['salary_per_day']}, salary_per_month - {employee_acceptance.cleaned_data['salary_per_month']})"
            Logs.objects.create(employee=employee.employee, action=action,
                                action_by=request.user, action_date=datetime.now())
            messages.success(
                request, f'Employee acceptance form details has been successfully saved.')
            return redirect('/employees/employee-acceptance/' + str(pk))
    else:
        # return HttpResponse("sfsd")
        employee_acceptance = EmployeeAcceptanceForm(instance=employee)

    context = {
        'head': 'Add Employee Acceptance Form Details',
        'form': employee_acceptance,
        'employee_id': pk,
        'pk': pk,
        'for_update': 1,
        'employee': employee.employee
    }
    return render(request, 'employee/acceptance.html', context)


@login_required
def employee_biodata(request, pk):
    employee = get_object_or_404(Employee, pk=pk)
    employee_hiring_details = get_object_or_404(
        Employee_hiring_details, employee=pk)
    employee_preferences = get_object_or_404(Employee_preferences, employee=pk)
    employment_record = Employment_record.objects.filter(employee=pk)
    employee_preferencesss = Employee_preferences.objects.filter(employee=pk)
    print(employee)
    response = HttpResponse(content_type='text/ms-excel')
    file_name = f"{employee} - biodata.xls"
    response['Content-Disposition'] = f'attachment; filename="{file_name}"'
    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('biodata')
    style = 'align: wrap on, vert centre, horiz center; font: bold on'
#    ws.row(0).write(0, value, xlwt.Style.easyxf(style))

    gen_settings = General_settings.objects.get(id=1)
    template_name = gen_settings.template_name
    company_name = gen_settings.main_company
    company_address = gen_settings.company_address
    company_contacts = gen_settings.company_contacts

    ws.write_merge(0, 1, 0, 9, template_name,
                   xlwt.Style.easyxf(style))
    style = 'align: wrap on, vert centre, horiz center;'
    ws.write_merge(2, 2, 0, 9, company_address,
                   xlwt.Style.easyxf(style))
    ws.write_merge(3, 3, 0, 9, company_address,
                   xlwt.Style.easyxf(style))
    style = 'align: wrap on, vert centre, horiz center; font: bold on'
    ws.write_merge(5, 5, 0, 1, 'PERSONAL DATA',
                   xlwt.Style.easyxf(style))
    style = 'align: wrap on, vert centre, horiz center;'
    ws.write_merge(7, 7, 0, 2, 'POSITION DESIRED:',
                   xlwt.Style.easyxf(style))

    ws.write_merge(7, 7, 3, 4, employee_hiring_details.position,
                   xlwt.Style.easyxf(style))

    ws.write_merge(7, 7, 7, 7, f'DATE:',
                   xlwt.Style.easyxf(style))

    ws.write_merge(7, 7, 8, 9, f'{datetime.now()}',
                   xlwt.Style.easyxf(style))

    ws.write_merge(8, 8, 0, 1, 'NAME:',
                   xlwt.Style.easyxf(style))

    ws.write_merge(8, 8, 2, 6, F'{employee.first_name} {employee.middle_name} {employee.last_name}',
                   xlwt.Style.easyxf(style))

    ws.write_merge(8, 8, 7, 7, f'SEX:',
                   xlwt.Style.easyxf(style))

    ws.write_merge(8, 8, 8, 9, f'{employee.gender}',
                   xlwt.Style.easyxf(style))

    ws.write_merge(9, 9, 0, 1, 'CITY ADDRESS:',
                   xlwt.Style.easyxf(style))

    ws.write_merge(9, 9, 2, 5, employee.address,
                   xlwt.Style.easyxf(style))

    ws.write_merge(10, 10, 0, 1, 'PROVINCIAL ADDRESS:',
                   xlwt.Style.easyxf(style))

    ws.write_merge(10, 10, 2, 5, employee.provincial_address,
                   xlwt.Style.easyxf(style))

    ws.write_merge(11, 11, 0, 1, 'BIRTHDAY:',
                   xlwt.Style.easyxf(style))

    ws.write_merge(11, 11, 2, 5, F'{employee.date_of_birth}',
                   xlwt.Style.easyxf(style))

    ws.write_merge(11, 11, 7, 7, f'PLACE:',
                   xlwt.Style.easyxf(style))

    ws.write_merge(11, 11, 8, 9, f'{employee.place_of_birth}',
                   xlwt.Style.easyxf(style))

    ws.write_merge(12, 12, 0, 1, 'CIVIL STATUS:',
                   xlwt.Style.easyxf(style))

    ws.write_merge(12, 12, 2, 5, F'{employee.civil_status}',
                   xlwt.Style.easyxf(style))

    ws.write_merge(12, 12, 6, 7, f'CITIZENSHIP:',
                   xlwt.Style.easyxf(style))

    ws.write_merge(12, 12, 8, 9, f'{employee.citizenship}',
                   xlwt.Style.easyxf(style))

    ws.write_merge(13, 13, 0, 1, 'Religion:',
                   xlwt.Style.easyxf(style))

    ws.write_merge(13, 13, 2, 5, F'',
                   xlwt.Style.easyxf(style))

    ws.write_merge(13, 13, 7, 7, f'PHONE:',
                   xlwt.Style.easyxf(style))

    ws.write_merge(13, 13, 8, 9, f'{employee.phone}',
                   xlwt.Style.easyxf(style))

    ws.write_merge(14, 14, 0, 1, 'SSS NO:',
                   xlwt.Style.easyxf(style))

    ws.write_merge(14, 14, 2, 4, F'{employee.sss_no}',
                   xlwt.Style.easyxf(style))

    ws.write_merge(14, 14, 5, 6, f'PHILHEALTH NO:',
                   xlwt.Style.easyxf(style))

    ws.write_merge(14, 14, 7, 9, f'{employee.philhealth_no}',
                   xlwt.Style.easyxf(style))

    ws.write_merge(15, 15, 0, 1, 'PAGIBIG NO:',
                   xlwt.Style.easyxf(style))

    ws.write_merge(15, 15, 2, 4, F'{employee.pagibig_no}',
                   xlwt.Style.easyxf(style))

    ws.write_merge(15, 15, 5, 6, f'TIN NO:',
                   xlwt.Style.easyxf(style))

    ws.write_merge(15, 15, 7, 9, f'{employee.tin_no}',
                   xlwt.Style.easyxf(style))

    ws.write_merge(16, 16, 0, 1, 'SPOUSE:',
                   xlwt.Style.easyxf(style))

    ws.write_merge(16, 16, 2, 4, F'{employee_preferences.spouse}',
                   xlwt.Style.easyxf(style))

    ws.write_merge(16, 16, 5, 6, f'OCCUPATION:',
                   xlwt.Style.easyxf(style))

    ws.write_merge(16, 16, 7, 9, f'{employee_preferences.occupation}',
                   xlwt.Style.easyxf(style))

    ws.write_merge(17, 17, 0, 1, 'ADDRESS:',
                   xlwt.Style.easyxf(style))

    ws.write_merge(17, 17, 2, 7, F'{employee_preferences.spouse_address}',
                   xlwt.Style.easyxf(style))

    ws.write_merge(18, 18, 0, 5, 'NO. OF CHILDREN, THEIR NAMES AND DATE OF BIRTH:',
                   xlwt.Style.easyxf(style))

    ws.write_merge(21, 21, 0, 1, "FATHER'S NAME:",
                   xlwt.Style.easyxf(style))

    ws.write_merge(21, 21, 2, 4, "",
                   xlwt.Style.easyxf(style))

    ws.write_merge(21, 21, 5, 6, "OCCUPATION:",
                   xlwt.Style.easyxf(style))

    ws.write_merge(21, 21, 7, 9, "",
                   xlwt.Style.easyxf(style))

    ws.write_merge(22, 22, 0, 1, "MOTHER'S NAME:",
                   xlwt.Style.easyxf(style))

    ws.write_merge(22, 22, 2, 4, "",
                   xlwt.Style.easyxf(style))

    ws.write_merge(22, 22, 5, 6, "OCCUPATION:",
                   xlwt.Style.easyxf(style))

    ws.write_merge(22, 22, 7, 9, "",
                   xlwt.Style.easyxf(style))

    ws.write_merge(23, 23, 0, 4, "PERSON TO CONTACT IN CASE OF EMERGENCY:",
                   xlwt.Style.easyxf(style))

    ws.write_merge(23, 23, 5, 9, "",
                   xlwt.Style.easyxf(style))

    ws.write_merge(24, 24, 0, 2, "HIS/HER ADDRESS AND TEL#:",
                   xlwt.Style.easyxf(style))

    ws.write_merge(24, 24, 3, 9, "",
                   xlwt.Style.easyxf(style))

    style = 'align: wrap on, vert centre, horiz center; font: bold on'

    ws.write_merge(26, 26, 0, 2, "EDUCATIONAL BACKGROUND",
                   xlwt.Style.easyxf(style))

    style = 'align: wrap on, vert centre, horiz center;'

    ws.write_merge(28, 28, 0, 1, "ELEMENTARY:",
                   xlwt.Style.easyxf(style))

    ws.write_merge(28, 28, 2, 5, "",
                   xlwt.Style.easyxf(style))

    ws.write_merge(28, 28, 6, 7, "DATE GRADUATED:",
                   xlwt.Style.easyxf(style))

    ws.write_merge(28, 28, 8, 9, "",
                   xlwt.Style.easyxf(style))

    ws.write_merge(29, 29, 0, 1, "HIGH SCHOOL:",
                   xlwt.Style.easyxf(style))

    ws.write_merge(29, 29, 2, 5, "",
                   xlwt.Style.easyxf(style))

    ws.write_merge(29, 29, 6, 7, "DATE GRADUATED:",
                   xlwt.Style.easyxf(style))

    ws.write_merge(29, 29, 8, 9, "",
                   xlwt.Style.easyxf(style))

    ws.write_merge(30, 30, 0, 1, "COLLEGE:",
                   xlwt.Style.easyxf(style))

    ws.write_merge(30, 30, 2, 5, "",
                   xlwt.Style.easyxf(style))

    ws.write_merge(30, 30, 6, 7, "DATE GRADUATED:",
                   xlwt.Style.easyxf(style))

    ws.write_merge(30, 30, 8, 9, "",
                   xlwt.Style.easyxf(style))

    ws.write_merge(31, 31, 0, 1, "COURSE:",
                   xlwt.Style.easyxf(style))

    ws.write_merge(31, 31, 2, 9, "",
                   xlwt.Style.easyxf(style))

    ws.write_merge(32, 32, 0, 1, "SPECIAL SKILLS:",
                   xlwt.Style.easyxf(style))

    ws.write_merge(32, 32, 2, 9, "",
                   xlwt.Style.easyxf(style))

    style = 'align: wrap on, vert centre, horiz center; font: bold on'

    ws.write_merge(34, 34, 0, 2, "EMPLOYMENT RECORDS",
                   xlwt.Style.easyxf(style))

    style = 'align: wrap on, vert centre, horiz center;'

    ws.write_merge(35, 35, 0, 2, "FROM",
                   xlwt.Style.easyxf(style))
    ws.write_merge(35, 35, 3, 5, "TO",
                   xlwt.Style.easyxf(style))
    ws.write_merge(35, 35, 6, 9, "COMPANY/POSITION",
                   xlwt.Style.easyxf(style))

    rowrow = 36

    if employment_record:
        for emp in employment_record:
            ws.write_merge(rowrow, rowrow, 0, 2, f"{emp.from_date}",
                           xlwt.Style.easyxf(style))
            ws.write_merge(rowrow, rowrow, 3, 5, f"{emp.to_date}",
                           xlwt.Style.easyxf(style))
            ws.write_merge(rowrow, rowrow, 6, 9, f"{emp.company}/{emp.position}",
                           xlwt.Style.easyxf(style))

            rowrow += 1
    style = 'align: wrap on, vert centre, horiz center; font: bold on'

    rowrow += 2

    ws.write_merge(rowrow, rowrow, 0, 2, "CHARACTER REFERENCES",
                   xlwt.Style.easyxf(style))

    style = 'align: wrap on, vert centre, horiz center;'
    rowrow += 1
    ws.write_merge(rowrow, rowrow, 0, 2, "NAME",
                   xlwt.Style.easyxf(style))
    ws.write_merge(rowrow, rowrow, 3, 5, "OCCUPATION",
                   xlwt.Style.easyxf(style))
    ws.write_merge(rowrow, rowrow, 6, 9, "ADDRESS",
                   xlwt.Style.easyxf(style))
    rowrow += 1
    if employee_preferencesss:
        for emp in employee_preferencesss:
            ws.write_merge(rowrow, rowrow, 0, 2, f"{emp.character_reference_name}",
                           xlwt.Style.easyxf(style))
            ws.write_merge(rowrow, rowrow, 3, 5, f"",
                           xlwt.Style.easyxf(style))
            ws.write_merge(rowrow, rowrow, 6, 9, f"{emp.character_reference_address}",
                           xlwt.Style.easyxf(style))
            rowrow += 1

    wb.save(response)

    Logs.objects.create(employee=employee, action=f"{employee} resume / biodata was successfully downloaded.",
                        action_by=request.user, action_date=datetime.now())

    return response

    # return HttpResponse(pk)


@login_required
def search_employees(request):
    company = request.POST.get("company")
    employee_name = request.POST.get("employee_name")
    payment_method = request.POST.get("payment_method")
    start_date = request.POST.get("start_date")
    end_date = request.POST.get("end_date")
    search_by = request.POST.get("search_by")
    order = request.POST.get("order")

    if start_date:
        start_date = start_date.replace("/", "-")
        start_date = list(start_date)
        start_date = start_date[6] + start_date[7] + start_date[8] + start_date[9] + \
            start_date[2] + start_date[0] + start_date[1] + \
            start_date[2] + start_date[3] + start_date[4]

    if end_date:
        end_date = end_date.replace("/", "-")
        end_date = list(end_date)
        end_date = end_date[6] + end_date[7] + end_date[8] + end_date[9] + \
            end_date[2] + end_date[0] + end_date[1] + \
            end_date[2] + end_date[3] + end_date[4]

    f1 = Employee_hiring_details.objects.filter(
        Q(employee__first_name__icontains=employee_name) |
        Q(employee__middle_name__icontains=employee_name) |
        Q(employee__last_name__icontains=employee_name)
    )
    f2 = f1.filter(payment_method__icontains=payment_method)
    f3 = f2.filter(employee__company__company_name__icontains=company)

    if search_by == "date_hired":
        f3 = Employee_hiring_details.objects.filter(
            employee__date_hired__range=[start_date, end_date])
    elif search_by == "contract_expiration":
        f3 = Employee_hiring_details.objects.filter(
            employee__contract_expiration__range=[start_date, end_date])

    if order == "ascending":
        f3 = f3.order_by("employee__last_name")
    elif order == "descending":
        f3 = f3.order_by("-employee__last_name")

    f3_count = f3.count()
    search_head = "Search result: 0"
    if f3_count > 0:
        search_head = "Search result: " + str(f3_count)
    context = {
        'head': search_head,
        'search_results': f3
    }

    return render(request, 'employee/search_results.html', context)

# Create your views here.
# employee list


@login_required
def employee_show(request):
    employees = Employee.objects.all()
    companies = Company.objects.all()
    context = {
        'title': 'company',
        'head': 'Employees',
        'employees': employees,
        'companies': companies
    }
    return render(request, 'employee/all_employees.html', context)


@login_required  # employee add
def employee_add(request):
    if request.method == 'POST':
        empAddForm = EmployeeAddForm(request.POST)
        company = get_object_or_404(Company, pk=request.POST['company'])
        if empAddForm.is_valid():
            citizenship = empAddForm.cleaned_data['citizenship']
            citizenship_option = request.POST['citizenship_option']
            # check if same names do exists
            names = Employee.objects.filter(
                first_name=empAddForm.cleaned_data['first_name'], middle_name=empAddForm.cleaned_data['middle_name'], last_name=empAddForm.cleaned_data['last_name']).first()
            if names:
                messages.error(
                    request, f"Employee {empAddForm.cleaned_data['first_name']} {empAddForm.cleaned_data['middle_name']} {empAddForm.cleaned_data['last_name']} already exists.")
                return redirect("employee-add")

            # return HttpResponse(request.POST['citizenship_option'])
            if citizenship == "Others":
                saved_form = empAddForm.save(commit=False)
                saved_form.citizenship = citizenship_option
                # add new citizenship to employee_citizenship
                citiz = Employee_citizenship()
                citiz.citizenship = citizenship_option
                citiz.save()
                saved_form.save()

            else:
                empAddForm.save()

            action = f"Employee {request.POST['first_name']} {request.POST['middle_name']} {request.POST['last_name']} was successfully added to {company} company."
            Logs.objects.create(
                action=action, action_by=request.user, action_date=datetime.now())

            messages.success(request, f'Employee has been successfully saved.')
            return redirect('employee-show')
    else:
        empAddForm = EmployeeAddForm(initial={'pagibig_no': "12"})

    context = {
        'head': 'Add Employee',
        'form': empAddForm
    }
    return render(request, 'employee/employee_add.html', context)

# employee update profile


@login_required
def employee_update(request, pk):
    employee = get_object_or_404(Employee, pk=pk)
    hiring_details = get_object_or_404(Employee_hiring_details, employee=pk)
    saved_company = employee.company
    saved_date_hired = employee.date_hired
    saved_contract_expiration = employee.contract_expiration
    saved_position = hiring_details.position
    if request.method == 'POST':
        form = EmployeeAddForm(request.POST or None, instance=employee)

        company_details = get_object_or_404(
            Company, pk=request.POST['company'])
        # return HttpResponse(f"{employee.company} - {company_details.company_name}")
        action = f"Employee {employee}  was successfully updated from\
                     'Company - {employee.company}, Firstname - {employee.first_name}, Middlename - {employee.middle_name}, Lastname - {employee.last_name},  id - {employee.emp_id}, \
                    address - {employee.address}, prov address - {employee.provincial_address}, bday - {employee.date_of_birth}, gender- {employee.gender}, place of birth - {employee.place_of_birth}, \
                    phone - {employee.phone}, date hired - {employee.date_hired}, contract expiration - {employee.contract_expiration}, gov deductions to implement - {employee.gov_deductions_to_implement}, \
                    sss option - {employee.sss_option}, sss bracket - {employee.sss_bracket}, sss value - {employee.sss_value}, pagbig value - {employee.pagibig_value}, philhealth value - {employee.philhealth_value}, \
                    sss no - {employee.sss_no}, pagibig no - {employee.pagibig_no}, philhealth no - {employee.philhealth_no}, tin no - {employee.tin_no}, civil status - {employee.civil_status}, citizenship - {employee.citizenship}, \
                    remarks - {employee.remarks},' \
                    to =============================================================>\
                    company - {company_details.company_name}, Firstname - {request.POST['first_name']}, Middlename - {request.POST['middle_name']}, Lastname - {request.POST['last_name']},  id - {request.POST['emp_id']}, \
                    address - {request.POST['address']}, prov address - {request.POST['provincial_address']}, bday - {request.POST['date_of_birth']}, gender- {request.POST['gender']}, place of birth - {request.POST['place_of_birth']}, \
                    phone - {request.POST['phone']}, date hired - {request.POST['date_hired']}, contract expiration - {request.POST['contract_expiration']}, gov deductions to implement - {request.POST['gov_deductions_to_implement']}, \
                    sss option - {request.POST['sss_option']}, sss bracket - {request.POST['sss_bracket']}, sss value - {request.POST['sss_value']}, pagbig value - {request.POST['pagibig_value']}, philhealth value - {request.POST['philhealth_value']}, \
                    sss no - {request.POST['sss_no']}, pagibig no - {request.POST['pagibig_no']}, philhealth no - {request.POST['philhealth_no']}, tin no - {request.POST['tin_no']}, civil status - {request.POST['civil_status']}, citizenship - {request.POST['citizenship']}, \
                    remarks - {request.POST['remarks']} "

        if form.is_valid():
            if saved_company != company_details.company_name:
                # not equal, save sa work record
                 # get hiring details

                Employment_record.objects.create(employee_id=pk, from_date=saved_date_hired,
                                                 to_date=saved_contract_expiration, company=saved_company, position=saved_position)
                Logs.objects.create(employee=employee, action=f"New work record was successfully added to {employee}. (from_date={saved_date_hired}, to_date={saved_contract_expiration}, company={saved_company}, position={saved_position}) ",
                                    action_by=request.user, action_date=datetime.now())

                # save employee update
                citizenship = form.cleaned_data['citizenship']
                citizenship_option = request.POST['citizenship_option']
                # return HttpResponse(request.POST['citizenship_option'])
                if citizenship == "Others":
                    saved_form = form.save(commit=False)
                    saved_form.citizenship = citizenship_option
                    # add new citizenship to employee_citizenship
                    citiz = Employee_citizenship()
                    citiz.citizenship = citizenship_option
                    citiz.save()
                    saved_form.save()
                    Logs.objects.create(employee=employee, action=action,
                                        action_by=request.user, action_date=datetime.now())
                else:
                    form.save()
                    Logs.objects.create(employee=employee, action=action,
                                        action_by=request.user, action_date=datetime.now())
                messages.success(request, 'Employee was successfully updated.')
                return redirect('employee-update', pk=pk)
            else:

                citizenship = form.cleaned_data['citizenship']
                citizenship_option = request.POST['citizenship_option']
                # return HttpResponse(request.POST['citizenship_option'])
                if citizenship == "Others":
                    saved_form = form.save(commit=False)
                    saved_form.citizenship = citizenship_option
                    # add new citizenship to employee_citizenship
                    citiz = Employee_citizenship()
                    citiz.citizenship = citizenship_option
                    citiz.save()
                    saved_form.save()

                    # action = f"Employee {request.POST['first_name']} {request.POST['middle_name']} {request.POST['last_name']} was successfully updated."
                    Logs.objects.create(employee=employee, action=action,
                                        action_by=request.user, action_date=datetime.now())

                else:
                    form.save()
                    # action = f"Employee {request.POST['first_name']} {request.POST['middle_name']} {request.POST['last_name']} was successfully updated."
                    Logs.objects.create(employee=employee, action=action,
                                        action_by=request.user, action_date=datetime.now())
                messages.success(request, 'Employee was successfully updated.')
                return redirect('employee-update', pk=pk)
    else:
        form = EmployeeAddForm(instance=employee)

    context = {
        'head': 'Update Employee',
        'form': form,
        'for_update': 1,
        'pk': pk,
                'employee': employee
    }
    return render(request, 'employee/employee_add.html', context)

# employee update preferences


@login_required
def employee_preferences(request, pk):
    employee = get_object_or_404(Employee_preferences, employee=pk)
    if request.method == 'POST':
        form = PreferencesForm(request.POST or None, instance=employee)
        if form.is_valid():
            form.save()
            action = f"{employee.employee} record preferences was successfully updated. (spouse - {form.cleaned_data['spouse']}, occupation - {employee.occupation}, spouse address - {employee.spouse_address}, character_reference_name- {employee.character_reference_name}, character_reference_address - {employee.character_reference_address}, character_reference_phone - {employee.character_reference_phone})"
            Logs.objects.create(employee=employee.employee, action=action,
                                action_by=request.user, action_date=datetime.now())

            messages.success(
                request, 'Employee preferences was successfully updated.')
            return redirect('employee-preferences', pk=pk)
    else:
        form = PreferencesForm(instance=employee)

    context = {
        'head': 'Update Employee - ' + employee.employee.first_name + ' ' + employee.employee.last_name,
        'form': form,
        'for_update': 1,
        'pk': pk,
                'employee': employee
    }
    return render(request, 'employee/employee_add.html', context)

# employee add employment history / records


@login_required
def update_record(request, record):
    employee_record = get_object_or_404(Employment_record, pk=record)
    if request.method == 'POST':
        form = UpdateRecord(request.POST or None, instance=employee_record)
        if form.is_valid():
            records = form.save(commit=False)
            records.employee = employee_record.employee
            records.save()

            action = f"Employee {employee_record.employee} (from date = {form.cleaned_data['from_date']}, to date = {form.cleaned_data['to_date']}, company = {form.cleaned_data['company']}, position = {form.cleaned_data['position']}) records was successfully updated."
            Logs.objects.create(
                action=action, action_by=request.user, action_date=datetime.now())

            messages.success(
                request, 'Employment record was successfully updated.')
            return redirect('employee-records', pk=employee_record.employee.pk)
    else:
        form = UpdateRecord(instance=employee_record)

    context = {
        'head': 'Update Employee Record- ' + employee_record.employee.first_name + ' ' + employee_record.employee.last_name,
        'form': form,
        'for_update': 1,
        'pk': employee_record.employee.pk,
        'employee': employee_record.employee
    }
    return render(request, 'employee/employee_add.html', context)


@login_required
def employee_records(request, pk):
    employee = get_object_or_404(Employee, pk=pk)
    may_records = Employment_record.objects.filter(employee=pk)
    if request.method == 'POST':
        form = RecordsForm(request.POST)
        if form.is_valid():
            records = form.save(commit=False)
            records.employee = employee
            records.save()
            # return HttpResponse(form)
            Logs.objects.create(employee=employee, action=f"New work record was successfully added to {employee}. (from_date={form.cleaned_data['from_date']}, to_date={form.cleaned_data['to_date']}, company={form.cleaned_data['company']}, position={form.cleaned_data['position']}) ",
                                action_by=request.user, action_date=datetime.now())

            messages.success(
                request, 'Employment record was successfully updated.')
            return redirect('employee-records', pk=pk)
    else:
        form = RecordsForm(instance=employee)

    context = {
        'head': 'Update Employee - ' + employee.first_name + ' ' + employee.last_name,
        'form': form,
        'for_update': 1,
        'pk': pk,
        'may_records': may_records,
                'employee': employee
    }
    return render(request, 'employee/employee_add.html', context)

# delete employment record


@login_required
def emp_record_delete(request, pk, emp_id):
    if request.method == 'POST':
        company = Employment_record.objects.filter(pk=pk).first()
        company.delete()
        messages.success(
            request, f'Employment record was successfully deleted.')
        return redirect('employee-records', pk=emp_id)
    return redirect('company-show')


# employee update hiring details

@login_required
def employee_hiring_details(request, pk):
    employee = get_object_or_404(Employee_hiring_details, employee=pk)
    employee_ = get_object_or_404(Employee, pk=pk)
    if request.method == 'POST':
        form = HiringDetailsForm(request.POST or None, instance=employee)
        if form.is_valid():
            form.save()
            action = f"{employee_} hiring details was successfully updated. (rate - {form.cleaned_data['rate']}, training_rate - {form.cleaned_data['training_rate']}, atm - {form.cleaned_data['atm']}, position - {form.cleaned_data['position']}, position - {form.cleaned_data['position']}, payment_method - {form.cleaned_data['payment_method']}, overtime multiplier - {form.cleaned_data['overtime_formula']})"
            Logs.objects.create(employee=employee_, action=action,
                                action_by=request.user, action_date=datetime.now())

            messages.success(
                request, 'Employee hiring details was successfully updated.')
            return redirect('employee-hiring-details', pk=pk)
    else:
        form = HiringDetailsForm(instance=employee)

    context = {
        'head': 'Update Employee Hiring Details - ' + employee.employee.first_name + ' ' + employee.employee.last_name,
        'form': form,
        'for_update': 1,
        'pk': pk,
                'employee': employee_
    }
    return render(request, 'employee/employee_add.html', context)

# employee update resume


@login_required
def employee_resume(request, pk):
    employee = get_object_or_404(Employee_resume, employee=pk)
    if request.method == 'POST':
        form = ResumeForm(request.POST or None,
                          request.FILES, instance=employee)
        # old_file = employee.resume
        if form.is_valid():
            # path_file = 'employee/media/' + str(old_file)
            # os.remove(os.path.join(settings.BASE_DIR, path_file))
            form.save()
            Logs.objects.create(employee=employee.employee, action="{employee.employee} was successfully downloaded.",
                                action_by=request.user, action_date=datetime.now())

            messages.success(
                request, 'Employee resume was successfully updated.')
            return redirect('employee-resume', pk=pk)
    else:
        form = ResumeForm(instance=employee)

    context = {
        'head': 'Update Employee - ' + employee.employee.first_name + ' ' + employee.employee.last_name,
        'form': form,
        'for_update': 1,
        'pk': pk,
        'employee': employee
    }
    return render(request, 'employee/employee_resume.html', context)


# employee sss contrib


@login_required
def employee_sss_contrib(request, pk):
    # employee = get_object_or_404(Employee_sss, employee=pk)
    employee_name = get_object_or_404(Employee, pk=pk)

    # check if its for phil asia
    if employee_name.company.company_name == "phil asia":
        employee = Payroll_for_phil_asia.objects.filter(employee=pk)
    else:
        employee = Payroll.objects.filter(employee=pk)

    if employee:
        context = {
            'head': 'SSS Contribution Employee - ' + employee_name.first_name + ' ' + employee_name.last_name,
            'for_update': 1,
            'pk': pk,
            'contribution': employee,
            'employee': employee_name,
            'kind': 'sss'
        }
        return render(request, 'employee/employee_gov_contrib.html', context)

    else:
        messages.error(request, f'No contributions yet.')
        return redirect('employee-update', pk=pk)


# employee pagibig contrib

@login_required
def employee_pagibig_contrib(request, pk):
    employee_name = get_object_or_404(Employee, pk=pk)

    # check if phil asia
    if employee_name.company.company_name == "phil asia":
        employee = Payroll_for_phil_asia.objects.filter(employee=pk)
    else:
        employee = Payroll.objects.filter(employee=pk)

    if employee:
        context = {
            'head': 'PAGIBIG Contribution Employee - ' + employee_name.first_name + ' ' + employee_name.last_name,
            'for_update': 1,
            'pk': pk,
            'contribution': employee,
            'employee': employee_name,
            'kind': 'pagibig'
        }
        return render(request, 'employee/employee_gov_contrib.html', context)

    else:
        messages.error(request, f'No contributions yet.')
        return redirect('employee-update', pk=pk)

# employee philhealth contrib


@login_required
def employee_philhealth_contrib(request, pk):
    employee_name = get_object_or_404(Employee, pk=pk)

    # check if phil asia
    if employee_name.company.company_name == "phil asia":
        employee = Payroll_for_phil_asia.objects.filter(employee=pk)
    else:
        employee = Payroll.objects.filter(employee=pk)

    if employee:
        context = {
            'head': 'PHILHEALTH Contribution Employee - ' + employee_name.first_name + ' ' + employee_name.last_name,
            'for_update': 1,
            'pk': pk,
            'contribution': employee,
            'employee': employee_name,
            'kind': 'philhealth'
        }
        return render(request, 'employee/employee_gov_contrib.html', context)

    else:
        messages.error(request, f'No contributions yet.')
        return redirect('employee-update', pk=pk)


# employee update company loan

@login_required
def employee_company_loan(request, pk):
    employee = get_object_or_404(Employee, pk=pk)
    may_loan = Employee_company_loan.objects.filter(employee=pk)
    if request.method == 'POST':
        form = CompanyLoanForm(request.POST)
        if form.is_valid():
            for loan in may_loan:
                if loan.status == False:
                    messages.error(request, 'There is still unpaid loan.')
                    return redirect('employee-company-loan', pk=pk)

            com_loan = form.save(commit=False)
            com_loan.employee = employee
            com_loan.save()
            action = f"{employee} added new loan. (loan amount - {form.cleaned_data['load_amount']}, rate to deduct - {form.cleaned_data['rate_to_deduct']})"
            Logs.objects.create(employee=employee, action=action,
                                action_by=request.user, action_date=datetime.now())
            messages.success(
                request, 'Employment company loan was successfully updated.')
            return redirect('employee-company-loan', pk=pk)
    else:
        form = CompanyLoanForm(instance=employee)
    context = {
        'head': 'Update Employee - ' + employee.first_name + ' ' + employee.last_name + " Company Loan",
        'form': form,
        'for_update': 1,
        'pk': pk,
        'may_loan': may_loan,
        'employee': employee,
        'kind': 'company'
    }
    return render(request, 'employee/employee_add.html', context)

# employee update company loan


@login_required
def employee_pagibig_loan(request, pk):
    employee = get_object_or_404(Employee, pk=pk)
    may_loan = Employee_pagibig_loan.objects.filter(employee=pk)
    if request.method == 'POST':
        form = PagibigLoanForm(request.POST)
        if form.is_valid():
            for loan in may_loan:
                if loan.status == False:
                    messages.error(request, 'There is still unpaid loan.')
                    return redirect('employee-pagibig-loan', pk=pk)

            com_loan = form.save(commit=False)
            com_loan.employee = employee
            com_loan.save()
            action = f"{employee} added new pagibig loan. (loan amount - {form.cleaned_data['load_amount']}, rate to deduct - {form.cleaned_data['rate_to_deduct']})"
            Logs.objects.create(employee=employee, action=action,
                                action_by=request.user, action_date=datetime.now())
            messages.success(
                request, 'Employment pagibig loan was successfully updated.')
            return redirect('employee-pagibig-loan', pk=pk)
    else:
        form = PagibigLoanForm(instance=employee)
    context = {
        'head': 'Update Employee - ' + employee.first_name + ' ' + employee.last_name + " Pagibig Loan",
        'form': form,
        'for_update': 1,
        'pk': pk,
        'may_loan': may_loan,
        'employee': employee,
        'kind': 'pagibig'
    }
    return render(request, 'employee/employee_add.html', context)


@login_required
def employee_sss_loan(request, pk):
    employee = get_object_or_404(Employee, pk=pk)
    may_loan = Employee_sss_loan.objects.filter(employee=pk)
    if request.method == 'POST':
        form = SssLoanForm(request.POST)
        if form.is_valid():
            for loan in may_loan:
                if loan.status == False:
                    messages.error(request, 'There is still unpaid loan.')
                    return redirect('employee-sss-loan', pk=pk)

            com_loan = form.save(commit=False)
            com_loan.employee = employee
            com_loan.save()
            action = f"{employee} added new sss loan. (loan amount - {form.cleaned_data['load_amount']}, rate to deduct - {form.cleaned_data['rate_to_deduct']})"
            Logs.objects.create(employee=employee, action=action,
                                action_by=request.user, action_date=datetime.now())
            messages.success(
                request, 'Employment sss loan was successfully updated.')
            return redirect('employee-sss-loan', pk=pk)
    else:
        form = SssLoanForm(instance=employee)
    context = {
        'head': 'Update Employee - ' + employee.first_name + ' ' + employee.last_name + " SSS Loan",
        'form': form,
        'for_update': 1,
        'pk': pk,
        'may_loan': may_loan,
        'employee': employee,
        'kind': 'sss'
    }
    return render(request, 'employee/employee_add.html', context)


# TODO: company load contrib
# employee company loan contrib

@login_required
def employee_comloan_contrib(request, pk, comloan):
    loans = Employee_comloan_contrib.objects.filter(company_loan=pk)
    comemployee = Employee.objects.filter(pk=comloan).first()
    employee = get_object_or_404(Employee, pk=comloan)
    if loans:
        context = {
            'head': 'Company Loan Contribution Employee - ' + comemployee.first_name + ' ' + comemployee.last_name,
            'for_update': 1,
            'pk': comloan,
            'contribution': loans,
            'employee': employee
        }
        return render(request, 'employee/employee_contrib.html', context)

    else:
        messages.error(request, f'No contributions yet.')
        return redirect('employee-company-loan', pk=comloan)

@login_required
def employee_valeloan_contrib(request, pk, comloan):
    # return HttpResponse(comloan)
    loans = Employee_valeloan_contrib.objects.filter(vale_loan=pk)
    comemployee = Employee.objects.filter(pk=comloan).first()
    employee = get_object_or_404(Employee, pk=comloan)
    if loans:
        context = {
            'head': 'Vale Loan Contribution Employee - ' + comemployee.first_name + ' ' + comemployee.last_name,
            'for_update': 1,
            'pk': comloan,
            'contribution': loans,
            'employee': employee
        }
        return render(request, 'employee/employee_contrib.html', context)

    else:
        messages.error(request, f'No contributions yet.')
        return redirect('employee-vale', pk=comloan)

def employee_canteen_contrib(request, pk, comloan):
    # return HttpResponse(comloan)
    loans = Employee_canteen_contrib.objects.filter(canteen_loan=pk)
    comemployee = Employee.objects.filter(pk=comloan).first()
    employee = get_object_or_404(Employee, pk=comloan)
    if loans:
        context = {
            'head': 'Canteen Loan Contribution Employee - ' + comemployee.first_name + ' ' + comemployee.last_name,
            'for_update': 1,
            'pk': comloan,
            'contribution': loans,
            'employee': employee
        }
        return render(request, 'employee/employee_contrib.html', context)

    else:
        messages.error(request, f'No contributions yet.')
        return redirect('employee-canteen', pk=comloan)

def employee_medical_contrib(request, pk, comloan):
    # return HttpResponse(comloan)
    loans = Employee_medical_contrib.objects.filter(medical_loan=pk)
    comemployee = Employee.objects.filter(pk=comloan).first()
    employee = get_object_or_404(Employee, pk=comloan)
    if loans:
        context = {
            'head': 'Medical Loan Contribution Employee - ' + comemployee.first_name + ' ' + comemployee.last_name,
            'for_update': 1,
            'pk': comloan,
            'contribution': loans,
            'employee': employee
        }
        return render(request, 'employee/employee_contrib.html', context)

    else:
        messages.error(request, f'No contributions yet.')
        return redirect('employee-medical', pk=comloan)


@login_required
def employee_pagibigloan_contrib(request, pk, comloan):
    loans = Employee_pagibigloan_contrib.objects.filter(pagibig_loan=pk)
    comemployee = Employee.objects.filter(pk=comloan).first()
    if loans:
        context = {
            'head': 'Pagibig Loan Contribution Employee - ' + comemployee.first_name + ' ' + comemployee.last_name,
            'for_update': 1,
            'pk': comloan,
            'contribution': loans,
            'employee': comemployee
        }
        return render(request, 'employee/employee_contrib.html', context)

    else:
        messages.error(request, f'No contributions yet.')
        return redirect('employee-pagibig-loan', pk=comloan)


@login_required
def employee_sssloan_contrib(request, pk, comloan):
    loans = Employee_sssloan_contrib.objects.filter(sss_loan=pk)
    comemployee = Employee.objects.filter(pk=comloan).first()
    if loans:
        context = {
            'head': 'SSS Loan Contribution Employee - ' + comemployee.first_name + ' ' + comemployee.last_name,
            'for_update': 1,
            'pk': comloan,
            'contribution': loans,
            'employee': comemployee
        }
        return render(request, 'employee/employee_contrib.html', context)

    else:
        messages.error(request, f'No contributions yet.')
        return redirect('employee-sss-loan', pk=comloan)


# employee update uniform

@login_required
def employee_uniform(request, pk):
    employee = get_object_or_404(Employee, pk=pk)
    may_uniform = Employee_uniform.objects.filter(employee=pk)
    if request.method == 'POST':
        form = UniformForm(request.POST)
        if form.is_valid():
            for uni in may_uniform:
                if uni.status == False:
                    messages.error(request, 'There is still unpaid Uniform.')
                    return redirect('employee-company-loan', pk=pk)

            uniform = form.save(commit=False)
            uniform.employee = employee
            uniform.save()
            messages.success(
                request, 'Uniform deduction was successfully updated.')
            return redirect('employee-uniform', pk=pk)
    else:
        form = UniformForm(instance=employee)
    context = {
        'head': 'Update Employee Uniform - ' + employee.first_name + ' ' + employee.last_name,
        'form': form,
        'for_update': 1,
        'pk': pk,
        'may_uniform': may_uniform,
        'employee': employee
    }
    return render(request, 'employee/employee_add.html', context)


# employee update medical

@login_required
def employee_medical(request, pk):
    employee = get_object_or_404(Employee, pk=pk)
    may_uniform = Employee_medical.objects.filter(employee=pk)
    if request.method == 'POST':
        form = MedicalForm(request.POST)
        if form.is_valid():
            for uni in may_uniform:
                if uni.status == False:
                    messages.error(request, 'There is still unpaid Medical.')
                    return redirect('employee-medical', pk=pk)

            uniform = form.save(commit=False)
            uniform.employee = employee
            uniform.save()
            action = f"{employee} added new medical. (amount - {form.cleaned_data['amount']}, rate to deduct - {form.cleaned_data['rate_to_deduct']})"
            Logs.objects.create(employee=employee, action=action,
                                action_by=request.user, action_date=datetime.now())
            messages.success(
                request, 'Medical deduction was successfully updated.')
            return redirect('employee-medical', pk=pk)
    else:
        form = MedicalForm(instance=employee)
    context = {
        'head': 'Update Employee Medical - ' + employee.first_name + ' ' + employee.last_name,
        'form': form,
        'for_update': 1,
        'pk': pk,
        'may_uniform': may_uniform,
        'employee': employee,
        'kind': 'medical'
    }
    return render(request, 'employee/employee_add.html', context)


# employee update canteen

@login_required
def employee_canteen(request, pk):
    employee = get_object_or_404(Employee, pk=pk)
    may_uniform = Employee_canteen.objects.filter(employee=pk)
    if request.method == 'POST':
        form = CanteenForm(request.POST)
        if form.is_valid():
            for uni in may_uniform:
                if uni.status == False:
                    messages.error(request, 'There is still unpaid canteen.')
                    return redirect('employee-canteen', pk=pk)

            uniform = form.save(commit=False)
            uniform.employee = employee
            uniform.save()
            action = f"{employee} added new canteen. (amount - {form.cleaned_data['amount']}, rate to deduct - {form.cleaned_data['rate_to_deduct']})"
            Logs.objects.create(employee=employee, action=action,
                                action_by=request.user, action_date=datetime.now())
            messages.success(
                request, 'Canteen deduction was successfully updated.')
            return redirect('employee-canteen', pk=pk)
    else:
        form = CanteenForm(instance=employee)
    context = {
        'head': 'Update Employee Canteen - ' + employee.first_name + ' ' + employee.last_name,
        'form': form,
        'for_update': 1,
        'pk': pk,
        'may_uniform': may_uniform,
        'employee': employee,
        'kind': 'canteen'
    }
    return render(request, 'employee/employee_add.html', context)


# employee update gatepass

@login_required
def employee_gatepass(request, pk):
    employee = get_object_or_404(Employee, pk=pk)
    may_uniform = Employee_gatepass.objects.filter(employee=pk)
    if request.method == 'POST':
        form = GatepassForm(request.POST)
        if form.is_valid():
            for uni in may_uniform:
                if uni.status == False:
                    messages.error(request, 'There is still unpaid gatepass.')
                    return redirect('employee-gatepass', pk=pk)

            uniform = form.save(commit=False)
            uniform.employee = employee
            uniform.save()
            action = f"{employee} added new gatepass. (amount - {form.cleaned_data['amount']}, rate to deduct - {form.cleaned_data['rate_to_deduct']})"
            Logs.objects.create(employee=employee, action=action,
                                action_by=request.user, action_date=datetime.now())
            messages.success(
                request, 'Gatepass deduction was successfully updated.')
            return redirect('employee-gatepass', pk=pk)
    else:
        form = GatepassForm(instance=employee)
    context = {
        'head': 'Update Employee Gatepass - ' + employee.first_name + ' ' + employee.last_name,
        'form': form,
        'for_update': 1,
        'pk': pk,
        'may_uniform': may_uniform,
        'employee': employee,
        'kind': 'gatepass'
    }
    return render(request, 'employee/employee_add.html', context)

def employee_gatepass_contrib(request, pk, comloan):
    # return HttpResponse(comloan)
    loans = Employee_gatepass_contrib.objects.filter(gatepass_loan=pk)
    comemployee = Employee.objects.filter(pk=comloan).first()
    employee = get_object_or_404(Employee, pk=comloan)
    if loans:
        context = {
            'head': 'Gatepass Loan Contribution Employee - ' + comemployee.first_name + ' ' + comemployee.last_name,
            'for_update': 1,
            'pk': comloan,
            'contribution': loans,
            'employee': employee
        }
        return render(request, 'employee/employee_contrib.html', context)

    else:
        messages.error(request, f'No contributions yet.')
        return redirect('employee-gatepass', pk=comloan)
# employee update vale

@login_required
def employee_vale(request, pk):
    employee = get_object_or_404(Employee, pk=pk)
    may_uniform = Employee_vale.objects.filter(employee=pk)
    if request.method == 'POST':
        form = ValeForm(request.POST)
        if form.is_valid():
            for uni in may_uniform:
                if uni.status == False:
                    messages.error(request, 'There is still unpaid vale.')
                    return redirect('employee-vale', pk=pk)

            uniform = form.save(commit=False)
            uniform.employee = employee
            uniform.save()
            action = f"{employee} added new vale. (amount - {form.cleaned_data['amount']}, rate to deduct - {form.cleaned_data['rate_to_deduct']})"
            Logs.objects.create(employee=employee, action=action,
                                action_by=request.user, action_date=datetime.now())
            messages.success(
                request, 'Vale deduction was successfully updated.')
            return redirect('employee-vale', pk=pk)
    else:
        form = ValeForm(instance=employee)
    context = {
        'head': 'Update Employee Vale - ' + employee.first_name + ' ' + employee.last_name,
        'form': form,
        'for_update': 1,
        'pk': pk,
        'may_uniform': may_uniform,
        'employee': employee,
        'vale': 'vale',
        'kind': 'vale'
    }
    return render(request, 'employee/employee_add.html', context)
