from django.urls import path

from . import views

urlpatterns = [
	path("", views.general_info, name="general-settings"),
	path("general_info", views.general_info, name="general-info"),
	path("sss_rates", views.sss_rates, name="sss-rates"),
	path("<int:id>/eeconrib_update/", views.sss_rates_update, name="eecontrib-update"),
	path("<int:id>/eeconrib_delete/", views.sss_rates_delete, name="eecontrib-delete"),
	path("eeconrib_create/", views.sss_rates_create, name="eecontrib-create"),
	path("bank_options", views.bank_options, name="bank-options"),

]