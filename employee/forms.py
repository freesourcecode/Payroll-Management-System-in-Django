from django import forms
from django.contrib.auth.models import User
from .models import Employee, Employee_preferences, Employment_record, Employee_hiring_details, Employee_resume, Employee_company_loan, Employee_uniform, Employee_medical, Employee_canteen, Employee_gatepass, Employee_vale, Employee_pagibig_loan, Employee_pagibigloan_contrib, Employee_sss_loan, Employee_sssloan_contrib, Employee_acceptance, Employee_return_to_work, Employee_leave_history, Employee_citizenship, Employee_picture, Employee_requirements, Employee_memo
from hrms.models import Company
from general_settings.models import BracketSSContribEE

class EmployeeAddForm(forms.ModelForm):
	date_of_birth = forms.DateField(widget=forms.DateInput(attrs={'class': 'datepicker'}))
	GENDER = CHOICES = (('', 'Choose gender'),('male', 'male'), ('female', 'female'))
	gender = forms.CharField(widget=forms.Select(choices=GENDER))

	# queryset= Employee_citizenship.objects.values_list('citizenship', flat=True).distinct()
	# CITIZENSHIP = [('', 'Choose citizenship')] + [(citizenship, citizenship) for citizenship in queryset]
	# citizenship = forms.CharField(widget=forms.Select(choices=CITIZENSHIP,attrs={'onchange': 'Hide()'}))

	def __init__(self, *args, **kwargs):
		super(EmployeeAddForm, self).__init__(*args, **kwargs)
		queryset= Employee_citizenship.objects.values_list('citizenship', flat=True).distinct()
		CITIZENSHIP = [('', 'Choose citizenship')] + [(citizenship, citizenship) for citizenship in queryset]
		self.fields['citizenship'] = forms.ChoiceField(choices=CITIZENSHIP,  widget=forms.Select(attrs={'onchange': 'Hide()'}))


	CIVIL_STATUS = CHOICES = (('', 'Choose status'),('single', 'single'), ('married', 'married'))
	civil_status = forms.CharField(widget=forms.Select(choices=CIVIL_STATUS))
	date_hired = forms.DateField(widget=forms.DateInput(attrs={'class': 'datepicker'}))
	contract_expiration = forms.DateField(widget=forms.DateInput(attrs={'class': 'datepicker'}))
	
	SSS_OPTION = CHOICES = (('0', 'Choose Sss Option'),
	  ('bracket', 'bracket'),
	  ('manual', 'manual'), )
	sss_option = forms.CharField(widget=forms.Select(choices=SSS_OPTION, attrs={'onchange': 'HideSSS()'}))
	

	USE_COM = CHOICES = (
	  ('', 'Choose Gov Deduction to Impelement'),
	  ('company_base_deductions', 'company_base_deductions'),
	  ('employee_base_deductions', 'employee_base_deductions'),
	)
	gov_deductions_to_implement = forms.CharField(widget=forms.Select(choices=USE_COM, attrs={'onchange': 'HideGovDeductions()'}))
	 
	brackets_queryset = BracketSSContribEE.objects.all()
	SSS_BRACKET = [('0', 'Choose Bracket')] + [(brackets.contrib_amount, brackets.ranged) for brackets in brackets_queryset]
	sss_bracket = forms.CharField(widget=forms.Select(choices=SSS_BRACKET))

	sss_no = forms.CharField(widget=forms.TextInput(attrs={'type':'number'}),)
	pagibig_no = forms.CharField(widget=forms.TextInput(attrs={'type':'number'}),)
	philhealth_no = forms.CharField(widget=forms.TextInput(attrs={'type':'number'}),)
	tin_no = forms.CharField(widget=forms.TextInput(attrs={'type':'number'}),)

	class Meta:
		model = Employee
		fields = ['first_name', 'middle_name', 'last_name', 'emp_id','address', 'provincial_address', 'date_of_birth', 'date_hired', 'contract_expiration','gender','place_of_birth', 'civil_status','phone', 'company', 'gov_deductions_to_implement','sss_option' ,'sss_bracket','sss_value','pagibig_value','philhealth_value',  'sss_no', 'pagibig_no', 'philhealth_no', 'tin_no','citizenship', 'remarks' ]

class PreferencesForm(forms.ModelForm):
	class Meta:
		model = Employee_preferences
		exclude = ['employee']

class RecordsForm(forms.ModelForm):
	from_date = forms.DateField(widget=forms.DateInput(attrs={'class': 'datepicker'}))
	to_date = forms.DateField(widget=forms.DateInput(attrs={'class': 'datepicker'}))
	class Meta:
		model = Employment_record
		exclude = ['employee']

class HiringDetailsForm(forms.ModelForm):
	PAYMENT_METHOD_CHOICES = CHOICES = (('', 'Choose method'),('weekly', 'weekly'), ('semi-monthly', 'semi-monthly'), ('monthly', 'monthly'))
	payment_method = forms.CharField(widget=forms.Select(choices=PAYMENT_METHOD_CHOICES))
	BANK_CHOICES = CHOICES = (('', 'Choose bank'),('BDO', 'BDO'), ('BPI', 'BPI'), ('AUB', 'AUB'))
	bank = forms.CharField(widget=forms.Select(choices=BANK_CHOICES))
	overtime_formula = forms.CharField(label="Input overtime multiplier. (ex. 1.25)")
	class Meta:
		model = Employee_hiring_details
		exclude = ['employee']

class ResumeForm(forms.ModelForm):
	class Meta:
		model = Employee_resume
		exclude = ['employee']

class PictureForm(forms.ModelForm):
	picture = forms.FileField(
        label='Select a picture',
        help_text=''
    )
	class Meta:
		model = Employee_picture
		exclude = ['employee']

class CompanyLoanForm(forms.ModelForm):
	load_amount = forms.CharField(label="Loan Amount")
	class Meta:
		model = Employee_company_loan
		exclude = ['employee', 'status']

class PagibigLoanForm(forms.ModelForm):
	class Meta:
		model = Employee_pagibig_loan
		exclude = ['employee', 'status']

class SssLoanForm(forms.ModelForm):
	class Meta:
		model = Employee_sss_loan
		exclude = ['employee', 'status']


class UniformForm(forms.ModelForm):
	class Meta:
		model = Employee_uniform
		exclude = ['employee', 'status']

class MedicalForm(forms.ModelForm):
	class Meta:
		exclude = ['employee', 'status']
		model = Employee_medical

class CanteenForm(forms.ModelForm):
	class Meta:
		model = Employee_canteen
		exclude = ['employee', 'status']

class GatepassForm(forms.ModelForm):
	class Meta:
		model = Employee_gatepass
		exclude = ['employee', 'status']


class ValeForm(forms.ModelForm):
	class Meta:
		model = Employee_vale
		exclude = ['employee', 'status']


class SearchForm(forms.ModelForm):
	company_name = forms.CharField(label='',
		  widget= forms.TextInput(attrs={'placeholder':'Company'}), required=False)

	class Meta:
		model = Company
		fields = ['company_name']

class EmployeeAcceptanceForm(forms.ModelForm):
	employment_status = CHOICES = (('', 'Choose status'),('Casual / Probationary', 'Casual / Probationary'), ('Seasonal', 'Seasonal'), ('Regular', 'Regular'))
	employment_status = forms.CharField(widget=forms.Select(choices=employment_status))
	start_date = forms.DateField(widget=forms.DateInput(attrs={'class': 'datepicker'}))
	end_date = forms.DateField(widget=forms.DateInput(attrs={'class': 'datepicker'}))
	class Meta:
		model = Employee_acceptance
		fields = ['employment_status', 'start_date', 'end_date', 'position', 'salary_per_day', 'salary_per_month']

class EmployeeReturnToWork(forms.ModelForm):
	return_on = forms.DateField(widget=forms.DateInput(attrs={'class': 'datepicker'}))
	absence_date_start = forms.DateField(widget=forms.DateInput(attrs={'class': 'datepicker'}))
	absence_date_end = forms.DateField(widget=forms.DateInput(attrs={'class': 'datepicker'}))
	class Meta:
		model = Employee_return_to_work
		fields = ['return_on', 'absence_date_start', 'absence_date_end', 'reason']


class EmployeeLeaveHistory(forms.ModelForm):
	start_date = forms.DateField(widget=forms.DateInput(attrs={'class': 'datepicker'}))
	end_date = forms.DateField(widget=forms.DateInput(attrs={'class': 'datepicker'}))
	class Meta:
		model = Employee_leave_history
		fields = ['start_date', 'end_date', 'no_of_days']

class UpdateRecord(forms.ModelForm):
	from_date = forms.DateField(widget=forms.DateInput(attrs={'class': 'datepicker'}))
	to_date = forms.DateField(widget=forms.DateInput(attrs={'class': 'datepicker'}))
	class Meta:
		model = Employment_record
		exclude = ['employee']

class Requirements(forms.ModelForm):
	class Meta:
		model = Employee_requirements
		exclude = ['employee']

class MemoForm(forms.ModelForm):
	class Meta:
		model = Employee_memo
		exclude = ['employee']