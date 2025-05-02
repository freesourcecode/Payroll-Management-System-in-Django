from django.urls import path
from .import views

urlpatterns = [
	path('payroll-list', views.payroll_list, name='payroll-list'),
	path('payroll-contributions', views.payroll_contributions, name='payroll-contributions'),
	path('payroll-csv/<int:pk>', views.payroll_csv, name='payroll-csv'),
	path('payroll-csv-phil-asia/<int:pk>', views.payroll_csv_phil_asia, name='payroll-csv-phil-asia'),
	path('payroll-payslip-phil-asia/<int:pk>', views.payroll_payslip_phil_asia, name='payroll-payslip-phil-asia'),
	path('payroll-billing-phil-asia/<int:pk>', views.payroll_billing_phil_asia, name='payroll-billing-phil-asia'),
	path('payroll-summary-phil-asia/<int:pk>', views.payroll_summary_phil_asia, name='payroll-summary-phil-asia'),
	path('payroll-billing/<int:pk>', views.payroll_billing, name='payroll-billing'),
	path('payroll-payslip/<int:pk>', views.payroll_payslip, name='payroll-payslip'),
	path('payroll-delete-base/<int:pk>', views.payroll_delete_base, name='payroll-delete-base'),
	path('payroll-encode/<int:payroll_view>/<int:pk>', views.payroll_encode, name='payroll-encode'),
	path('payroll-view/<int:pk>', views.payroll_view, name='payroll-view'),
	path('payroll-view-phil-asia/<int:pk>', views.payroll_view_phil_asia, name='payroll-view-phil-asia'),
	path('get-overtime-formula', views.get_overtime_formula, name='get-overtime-formula'),
	

	

	path('gov-benefits/', views.gov_benefits, name='gov-benefits')
]
