from django import forms
from django.contrib.auth.models import User
from .models import Company, Company_rates

class CompanyAddForm(forms.ModelForm):
	class Meta:
		model = Company
		fields = ['company_name', 'contact_person', 'phone', 'fax', 'email','status']

class CompanyRates(forms.ModelForm):
	class Meta:
		model = Company_rates
		fields = ['base_rate', 'activate_rates', 'base_training_rate', 'activate_training_rate']


class CompanyGovDeduct(forms.ModelForm):
	 class Meta:
	 	model = Company_rates
	 	fields = ['sss', 'pagibig', 'philhealth']
	 	labels = {
	 		'sss':'SSS',
	 		'pagibig':'Pagibig',
	 		'philhealth':'Philhealth',
	 	}


class CompanyOtherOptions(forms.ModelForm):
	class Meta:
		model = Company_rates
		fields = [
				'ecola_rate','activate_ecola', 'activate_overtime', 'activate_1_25_overtime', 'activate_holiday','activate_rest_day', 'activate_special',
				 'activate_night_differential', 'activate_thirteenth_month',
				'activate_sil', 'activate_tshirt', 'activate_rf', 'activate_house', 'activate_misc','activate_gatepass', 'activate_medical' , 'activate_pants',
				'activate_vale', 'activate_uniform' , 'activate_company_loan', 'activate_pagibig_loan' , 'activate_sss_loan', 'activate_canteen', 'activate_service_fee',
				'activate_sunday', 'activate_tardiness', 'activate_allowance', 'activate_adjustment', 'activate_transpo_allowance'
		]
