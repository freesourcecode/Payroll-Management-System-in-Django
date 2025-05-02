from django.shortcuts import render, HttpResponse, get_object_or_404, redirect
from .models import General_settings, BracketSSContribEE
from .forms import GeneralInfoForm, BracketSSContribEEForm
from django.contrib import messages


# def general_settings(request):
# 	context = {
# 		"head" : "General Settings",
# 	}
# 	return render(request, "general_settings/gs_home.html", context)


def general_info(request):
	data = get_object_or_404(General_settings, pk=1)
	form = GeneralInfoForm(request.POST or None, instance=data)

	if form.is_valid():
		form.save()
		messages.success(request, "Company info was successfully updated.")
		request.session['main_company'] = data.main_company
		return redirect('general-info')

	context = {
		"page_nick": 'general-info',
		"head" : "General Settings",
		"form": form,
		"main_company": data
	}
	return render(request, "general_settings/gs_home.html", context)

def sss_rates(request):
	# data = get_object_or_404(BracketSSContribEE, pk=1)
	# form = BracketSSContribEEForm(request.POST or None, instance=data)
	ee_sss_contrib = BracketSSContribEE.objects.all()

	context = {
		"page_nick":"sss-rates",
		"head" : "EE SSS Conrib",
		"ee_sss_contrib": ee_sss_contrib
	}
	return render(request, "general_settings/gs_eecontribss.html", context)

def sss_rates_create(request):
	form = BracketSSContribEEForm(request.POST or None)
	
	if form.is_valid():
		form.save()
		form = BracketSSContribEEForm()
		return redirect('sss-rates')

	context = {
		'page_nick':'sss-rates',
		'form': form
	}
	return render(request, 'general_settings/gs_home.html', context)


def sss_rates_delete(request, id):
	obj = get_object_or_404(BracketSSContribEE, id=id)
	obj.delete()
	return redirect('sss-rates')

def sss_rates_update(request, id):
	obj = get_object_or_404(BracketSSContribEE, id=id)
	form = BracketSSContribEEForm(request.POST or None, instance=obj)
	
	if form.is_valid():
		form.save()
		return redirect('sss-rates')

	context = {
		'page_nick':'sss-rates',
		'form': form
	}
	return render(request, 'general_settings/gs_home.html', context)


def bank_options(request):
	return HttpResponse('bank options')
