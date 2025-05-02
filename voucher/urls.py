from django.urls import path
from .import views

urlpatterns = [
    path('voucher', views.voucher, name='voucher'),
    path('voucher-add', views.voucher_add, name='voucher-add'),
    path('voucher-update/<int:pk>', views.voucher_update, name='voucher-update'),
    path('encode-particulars/<int:voucher_id>', views.encode_particulars, name='encode-particulars'),
    path('delete-particulars/<int:particular_id>', views.delete_particulars, name='delete-particulars'),
    path('download-voucher/<int:voucher_id>', views.download_voucher, name='download-voucher'),
]
