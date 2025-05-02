from django.db import models

from employee.models import Employee
from hrms.models import Company
# Create your models here.

class Base_payroll(models.Model):
    company = models.ForeignKey(Company, on_delete=models.CASCADE, related_name='company_payroll')
    start_date = models.DateField()
    end_date = models.DateField()
    payment_method = models.CharField(max_length=50)
    # last_cut_off_for_this_month = models.BooleanField(default=False)
    activate_gov_deductions = models.BooleanField(default=False)
    activate_company_loan_deductions = models.BooleanField(default=False)

    def __str__(self):
        return f'{self.company} - {self.start_date} - {self.end_date} - {self.payment_method}'


class Payroll_for_phil_asia(models.Model):
    base_payroll = models.ForeignKey(Base_payroll, on_delete=models.CASCADE, related_name='base_payroll_phil_asia')
    employee = models.ForeignKey(Employee, on_delete=models.CASCADE, related_name='payroll_employee_phil_asia')
    rate = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    training_rate = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    days = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    amount = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    hrs = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    pay = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    nd = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    nd_pay = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    hrs_ot = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    hrs_ot_pay = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    total_amount = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    canteen = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    office = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    sss = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    pagibig = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    philhealth = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    net_amount = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)



class Payroll(models.Model):
    base_payroll = models.ForeignKey(Base_payroll, on_delete=models.CASCADE, related_name='base_payroll')
    employee = models.ForeignKey(Employee, on_delete=models.CASCADE, related_name='payroll_employee')
    regular_days = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    rate = models.DecimalField(max_digits=7, decimal_places=2)
    regular_amount = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    training_days = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    training_rate = models.DecimalField(max_digits=7, decimal_places=2)
    training_amount = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    ecola = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    rest_days = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    rest_amount = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    rest_day_overtime = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    rest_day_overtime_amount = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    holiday_regular_days = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    holiday_regular_amount = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    holiday_overtime = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    holiday_overtime_amount = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    holiday_training_days = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    holiday_training_amount = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    overtime_regular = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    overtime_regular_amount = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    overtime_training_days = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    overtime_training_amount = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    special_holiday_days = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    special_holiday_amount = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    special_holiday_overtime = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    special_holiday_overtime_amount = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    sunday = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    sunday_amount = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    sunday_overtime = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    sunday_overtime_amount = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    sunday_nd = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    sunday_nd_amount = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    night_diff_days = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    night_diff_amount = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    gross = models.DecimalField(max_digits=7, decimal_places=2, null=True)
    tardiness_undertime_regular = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    tardiness_undertime_regular_amount = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    thirteenth_month = models.DecimalField(max_digits=7, decimal_places=2, default=0.0, blank=True)
    tshirt = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    rf = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    house = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    misc = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    sil = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    gatepass = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    uniform = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    medical = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    canteen = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    vale = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    pants = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    service_fee = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    company_loan = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    sss_loan = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    pagibig_loan = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    sss = models.DecimalField(max_digits=7, decimal_places=2, null=True, default=0.0)
    pagibig = models.DecimalField(max_digits=7, decimal_places=2, null=True, default=0.0)
    philhealth = models.DecimalField(max_digits=7, decimal_places=2, null=True, default=0.0)
    net_amount = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    sss_employer = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    pagibig_employer = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    philhealth_employer = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    allowance = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    adjustment = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    transpo_allowance = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    gross = models.DecimalField(max_digits=7, decimal_places=2, default=0.0)
    valid_for_deduct_company_loan = models.IntegerField(default=0)
    valid_for_deduct_vale = models.IntegerField(default=0)
    valid_for_deduct_canteen = models.IntegerField(default=0)
    valid_for_deduct_medical = models.IntegerField(default=0)
    valid_for_deduct_gatepass = models.IntegerField(default=0)
    valid_for_deduct_sss_loan = models.IntegerField(default=0)
    valid_for_deduct_pagibig_loan = models.IntegerField(default=0)
    remarks = models.IntegerField(default=False)

    def __str__(self):
        return f'{self.employee} payroll'


class Gov_benefits(models.Model):
    employee = models.ForeignKey(
        Employee, on_delete=models.CASCADE, related_name='employee_ben')
    days = models.DecimalField(max_digits=7, decimal_places=2)
    deductions = models.DecimalField(max_digits=7, decimal_places=2)
    start_date = models.DateField()
    end_date = models.DateField()
    def __str__(self):
        return f'{ self.employee } - benefits'
