from django import forms
from django.contrib.auth.models import User
from .models import Payroll, Base_payroll, Payroll_for_phil_asia
from django.forms import modelformset_factory

PayrollFullAddFormPhilAsia = modelformset_factory(
    Payroll_for_phil_asia,
    fields=[
        'employee','rate','days', 'amount', 'hrs', 'pay', 'nd', 'nd_pay', 'hrs_ot', 'hrs_ot_pay', 'total_amount', 'canteen', 'office', 'sss', 'pagibig', 'philhealth','net_amount'
    ],
    extra=0,
    widgets={
        'days': forms.TextInput(attrs={'onchange': 'get_regular_amount()', 'size': '20'}),
        'amount': forms.TextInput(attrs={'readonly': True, 'size': '20'}),
        'hrs': forms.TextInput(attrs={'onchange': 'get_hrs_pay()', 'size': '20'}),
        'pay': forms.TextInput(attrs={'readonly': True, 'size': '20'}),
        'nd': forms.TextInput(attrs={'onchange': 'get_nd_pay()', 'size': '20'}),
        'nd_pay': forms.TextInput(attrs={'readonly': True, 'size': '20'}),
        'hrs_ot': forms.TextInput(attrs={'onchange': 'get_ot_pay()', 'size': '20'}),
        'hrs_ot_pay': forms.TextInput(attrs={'readonly': True, 'size': '20'}),
        'canteen': forms.TextInput(attrs={'onchange': 'get_net()', 'size': '20'}),
        'office': forms.TextInput(attrs={'onchange': 'get_net()', 'size': '20'}),
        'sss': forms.TextInput(attrs={'onchange': 'get_net()', 'size': '20'}),
        'pagibig': forms.TextInput(attrs={'onchange': 'get_net()', 'size': '20'}),
        'philhealth': forms.TextInput(attrs={'onchange': 'get_net()', 'size': '20'}),
        'net_amount': forms.TextInput(attrs={'readonly': True, 'size': '20'}),

    }
)


PayrollFullAddForm = modelformset_factory(
    Payroll,
    fields=[
        'employee', 'regular_days', 'rate', 'regular_amount', 'ecola', 'overtime_regular', 'overtime_regular_amount',
        'sunday', 'sunday_amount', 'sunday_overtime', 'sunday_overtime_amount', 'sunday_nd', 'sunday_nd_amount',
        'holiday_regular_days', 'holiday_regular_amount',
        'holiday_overtime', 'holiday_overtime_amount',
        'special_holiday_days', 'special_holiday_amount',
        'special_holiday_overtime', 'special_holiday_overtime_amount',
        'rest_days', 'rest_amount', 'rest_day_overtime', 'rest_day_overtime_amount', 'night_diff_days', 'night_diff_amount',
        'tardiness_undertime_regular', 'tardiness_undertime_regular_amount', 'uniform', 'medical', 'canteen',
        'vale', 'pants', 'service_fee', 'thirteenth_month', 'sil', 'tshirt', 'rf', 'house', 'misc', 'gatepass',
        'company_loan', 'sss_loan', 'pagibig_loan', 'allowance', 'adjustment', 'transpo_allowance','sss', 'pagibig', 'philhealth','sss_employer', 'pagibig_employer', 'philhealth_employer', 'gross',
        'net_amount', 'valid_for_deduct_company_loan', 'valid_for_deduct_vale', 'valid_for_deduct_canteen', 'valid_for_deduct_medical', 'valid_for_deduct_gatepass', 'valid_for_deduct_sss_loan', 'valid_for_deduct_pagibig_loan'
                        # 'training_days' , 'training_rate', 'training_amount'
    ],
    extra=0,
    widgets={
        'regular_days': forms.TextInput(attrs={'onchange': 'get_regular_amount()', 'size': '20'}),
        # 'training_days': forms.TextInput(attrs={'onchange': 'get_training_amount()'}),
        'regular_amount': forms.TextInput(attrs={'readonly': True, 'size': '20'}),
        # 'training_amount': forms.TextInput(attrs={'id': 'training_amount'}),
        'rate': forms.TextInput(attrs={'readonly': True, 'size': '10'}),
        # 'training_rate': forms.TextInput(attrs={'readonly': True}),
        'ecola': forms.TextInput(attrs={'readonly': True}),
        'overtime_regular': forms.TextInput(attrs={'onchange': 'get_overtime_regular_amount()'}),
        'overtime_regular_amount': forms.TextInput(attrs={'readonly': True}),
        'sunday': forms.TextInput(attrs={'onchange': 'get_sunday_amount()'}),
        'sunday_amount': forms.TextInput(attrs={'readonly': True}),
        'sunday_overtime': forms.TextInput(attrs={'onchange': 'get_sunday_overtime_amount()'}),
        'sunday_overtime_amount': forms.TextInput(attrs={'readonly': True}),
        'sunday_nd': forms.TextInput(attrs={'onchange': 'get_sunday_nd_amount()'}),
        'sunday_nd_amount': forms.TextInput(attrs={'readonly': True}),

        'holiday_regular_days': forms.TextInput(attrs={'onchange': 'get_holiday_regular_amount()'}),
        'holiday_regular_amount': forms.TextInput(attrs={'readonly': True}),
        'holiday_overtime': forms.TextInput(attrs={'onchange': 'get_holiday_overtime_amount()'}),
        'holiday_overtime_amount': forms.TextInput(attrs={'readonly': True}),

        'special_holiday_days': forms.TextInput(attrs={'onchange': 'get_special_holiday_amount()'}),
        'special_holiday_amount': forms.TextInput(attrs={'readonly': True}),
        'special_holiday_overtime': forms.TextInput(attrs={'onchange': 'get_special_holiday_overtime_amount()'}),
        'special_holiday_overtime_amount': forms.TextInput(attrs={'readonly': True}),
        'rest_days': forms.TextInput(attrs={'onchange': 'get_rest_day_amount()'}),
        'rest_amount': forms.TextInput(attrs={'readonly': True}),
        'rest_day_overtime': forms.TextInput(attrs={'onchange': 'get_rest_day_overtime_amount()'}),
        'rest_day_overtime_amount': forms.TextInput(attrs={'readonly': True}),
        'night_diff_days': forms.TextInput(attrs={'onchange': 'get_night_diff_amount()'}),
        'night_diff_amount': forms.TextInput(attrs={'readonly': True}),
        'tardiness_undertime_regular': forms.TextInput(attrs={'onchange': 'get_tardiness_undertime_regular_amount()'}),
        'uniform': forms.TextInput(attrs={'onchange': 'get_uniform()'}),
        'medical': forms.TextInput(attrs={'onchange': 'get_medical()'}),
        'canteen': forms.TextInput(attrs={'onchange': 'get_canteen()'}),
        'tardiness_undertime_regular_amount': forms.TextInput(attrs={'readonly': True}),
        'gatepass': forms.TextInput(attrs={'onchange': 'get_gatepass()'}),
        'vale': forms.TextInput(attrs={'onchange': 'get_vale()'}),
        'thirteenth_month': forms.TextInput(attrs={'readonly': True}),
        'sil': forms.TextInput(attrs={'onchange': 'get_sil()'}),
        'tshirt': forms.TextInput(attrs={'onchange': 'get_tshirt()'}),
        'rf': forms.TextInput(attrs={'onchange': 'get_rf()'}),
        'house': forms.TextInput(attrs={'onchange': 'get_house()'}),
        'misc': forms.TextInput(attrs={'onchange': 'get_misc()'}),
        'pants': forms.TextInput(attrs={'onchange': 'get_pants()'}),
        'service_fee': forms.TextInput(attrs={'onchange': 'get_service_fee()'}),
        'company_loan': forms.TextInput(attrs={'readonly': True}),
        'sss_loan': forms.TextInput(attrs={'readonly': True}),
        'pagibig_loan': forms.TextInput(attrs={'readonly': True}),
        'allowance': forms.TextInput(attrs={'onchange': 'add_allowance()'}),
        'adjustment': forms.TextInput(attrs={'onchange': 'add_adjustment()'}),
        'transpo_allowance': forms.TextInput(attrs={'onchange': 'add_transpo_allowance()'}),

        'sss': forms.TextInput(attrs={'readonly': True}),
        'pagibig': forms.TextInput(attrs={'readonly': True}),
        'philhealth': forms.TextInput(attrs={'readonly': True}),
        'net_amount': forms.TextInput(attrs={'readonly': True}),
        'valid_for_deduct_company_loan': forms.HiddenInput(),
        'valid_for_deduct_vale': forms.HiddenInput(),
        'valid_for_deduct_canteen': forms.HiddenInput(),
        'valid_for_deduct_medical': forms.HiddenInput(),
        'valid_for_deduct_gatepass': forms.HiddenInput(),
        'valid_for_deduct_sss_loan': forms.HiddenInput(),
        'valid_for_deduct_pagibig_loan': forms.HiddenInput()
    }
)


class ContributionForm(forms.ModelForm):
    start_date = forms.DateField(widget=forms.DateInput(attrs={'class': 'datepicker'}))
    end_date = forms.DateField(widget=forms.DateInput(attrs={'class': 'datepicker'}))
    PAYMENT_METHOD = CHOICES = (('', 'Choose method'), ('weekly', 'weekly'),
                                ('semi-monthly', 'semi-monthly'), ('monthly', 'monthly'))
    payment_method = forms.CharField(widget=forms.Select(choices=PAYMENT_METHOD))
    CONTRIBUTION_KIND = CHOICES = (('', 'Choose kind'), ('vale', 'vale'), ('company_loan', 'company_loan'), (
        'sss_loan', 'sss_loan'), ('pagibig_loan', 'pagibig_loan'), ('sss', 'sss'), ('pagibig', 'pagibig'), ('philhealth', 'philhealth'))
    contribution_kind = forms.CharField(widget=forms.Select(choices=CONTRIBUTION_KIND))

    class Meta:
        model = Base_payroll
        fields = ['company', 'start_date', 'end_date', 'payment_method', 'contribution_kind']


class BasePayrollForm(forms.ModelForm):
    start_date = forms.DateField(widget=forms.DateInput(attrs={'class': 'datepicker'}))
    end_date = forms.DateField(widget=forms.DateInput(attrs={'class': 'datepicker'}))
    PAYMENT_METHOD = CHOICES = (('', 'Choose method'), ('weekly', 'weekly'),
                                ('semi-monthly', 'semi-monthly'), ('monthly', 'monthly'))
    payment_method = forms.CharField(widget=forms.Select(choices=PAYMENT_METHOD))
    PAYMENT_METHOD = CHOICES = (('', 'Choose method'), ('weekly', 'weekly'),
                                ('semi-monthly', 'semi-monthly'), ('monthly', 'monthly'))

    class Meta:
        model = Base_payroll
        fields = ['company', 'start_date', 'end_date', 'payment_method', 'activate_gov_deductions', 'activate_company_loan_deductions']


# class PayrollFullAddForm(forms.ModelForm):
#
#     class Meta:
#         model = Payroll
#         overtime_regular = forms.DecimalField(max_digits = 7, decimal_places = 2)
#         fields = [
#                     'regular_days', 'rate', 'regular_amount', 'ecola','overtime_regular', 'overtime_regular_amount',
#                     'special_holiday_days', 'special_holiday_amount',
#                     'special_holiday_overtime', 'special_holiday_overtime_amount',
#                     'rest_days', 'rest_amount','rest_day_overtime', 'rest_day_overtime_amount',
#                     'tardiness_undertime_regular', 'tardiness_undertime_regular_amount','uniform', 'medical', 'canteen',
#                     'vale', 'pants', 'thirteenth_month', 'sil', 'gatepass',
#                     'sss', 'pagibig', 'philhealth',
#                     'net_amount'
#                     # 'training_days' , 'training_rate', 'training_amount'
#                 ]
#         widgets = {
#             'regular_days': forms.TextInput(attrs={'onchange': 'get_regular_amount()'}),
#             # 'training_days': forms.TextInput(attrs={'onchange': 'get_training_amount()'}),
#             'regular_amount': forms.TextInput(attrs={'id': 'regular_amount'}),
#             # 'training_amount': forms.TextInput(attrs={'id': 'training_amount'}),
#             'rate': forms.TextInput(attrs={'readonly': True}),
#             # 'training_rate': forms.TextInput(attrs={'readonly': True}),
#             'ecola': forms.TextInput(attrs={'readonly': True}),
#             'overtime_regular': forms.TextInput(attrs={'onchange': 'get_overtime_regular_amount()'}),
#             'overtime_regular_amount': forms.TextInput(attrs={'readonly': True, 'id': 'overtime_regular_amount'}),
#             'special_holiday_days': forms.TextInput(attrs={'onchange': 'get_special_holiday_amount()'}),
#             'special_holiday_amount': forms.TextInput(attrs={'readonly': True, 'id': 'special_holiday_amount'}),
#             'special_holiday_overtime': forms.TextInput(attrs={'onchange': 'get_special_holiday_overtime_amount()'}),
#             'special_holiday_overtime_amount': forms.TextInput(attrs={'readonly': True, 'id': 'special_holiday_overtime_amount'}),
#             'rest_days': forms.TextInput(attrs={'onchange': 'get_rest_day_amount()'}),
#             'rest_amount': forms.TextInput(attrs={'readonly': True, 'id': 'rest_amount'}),
#             'rest_day_overtime': forms.TextInput(attrs={'onchange': 'get_rest_day_overtime_amount()'}),
#             'rest_day_overtime_amount': forms.TextInput(attrs={'readonly': True, 'id': 'rest_day_overtime_amount'}),
#             'tardiness_undertime_regular': forms.TextInput(attrs={'onchange': 'get_tardiness_undertime_regular_amount()'}),
#             'uniform': forms.TextInput(attrs={'onchange': 'get_uniform()'}),
#             'medical': forms.TextInput(attrs={'onchange': 'get_medical()'}),
#             'canteen': forms.TextInput(attrs={'onchange': 'get_canteen()'}),
#             'tardiness_undertime_regular_amount': forms.TextInput(attrs={'readonly': True, 'id': 'tardiness_undertime_regular_amount'}),
#             'gatepass': forms.TextInput(attrs={'onchange': 'get_gatepass()'}),
#             'vale': forms.TextInput(attrs={'onchange': 'get_vale()'}),
#             'thirteenth_month': forms.TextInput(attrs={'onchange': 'get_thirteenth_month()'}),
#             'sil': forms.TextInput(attrs={'onchange': 'get_sil()'}),
#             'pants': forms.TextInput(attrs={'onchange': 'get_pants()'}),
#             'sss': forms.TextInput(attrs={'readonly': True}),
#             'pagibig': forms.TextInput(attrs={'readonly': True}),
#             'philhealth': forms.TextInput(attrs={'readonly': True}),
#             'net_amount': forms.TextInput(attrs={'readonly': True , 'id': 'net_amount'})
#         }
