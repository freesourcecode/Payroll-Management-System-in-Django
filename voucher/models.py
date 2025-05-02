from django.db import models

class Voucher(models.Model):
	rc_no = models.IntegerField()
	date_created = models.DateField()
	place = models.CharField(max_length=50)
	voucher_no = models.CharField(max_length=50)
	voucher_created_date = models.DateField()
	paid_to = models.CharField(max_length=255)
	address = models.CharField(max_length=255)
    

class Voucher_particulars(models.Model):
	voucher = models.ForeignKey(
        Voucher, on_delete=models.CASCADE, related_name='voucher')
	particular_name = models.CharField(max_length=255)
	amount = models.DecimalField(max_digits=10, decimal_places=1)
		    
    
