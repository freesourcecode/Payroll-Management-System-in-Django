from django import forms
from django.contrib.auth.models import User
from .models import Voucher, Voucher_particulars

class VoucherForm(forms.ModelForm):
	date_created = forms.DateField(widget=forms.DateInput(attrs={'class': 'datepicker'}))
	voucher_created_date = forms.DateField(widget=forms.DateInput(attrs={'class': 'datepicker'}))
	
	class Meta:
		model = Voucher
		fields = ['rc_no', 'date_created', 'place', 'voucher_no', 'voucher_created_date', 'paid_to', 'address']

class ParticularsForm(forms.ModelForm):

	class Meta:
		model = Voucher_particulars
		fields = ['particular_name', 'amount']
