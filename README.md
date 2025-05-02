# Payroll Management System Project in Django with Source Code

A **Payroll Management System in Django**, keeps track of all of the employee’s information and data.

We’ve created all of the **employee’s**, **company**, **voucher**, **payrolls**, **logs**, and **general settings crud** (**create**, **read**, **update**, and **delete**) operations.

This is a role-based module in which the admin can perform any operation on the data.

The **payroll module** is typically used to **handle employees salary** while the company module handles company features.

The **Payroll Management System** is an easy project for beginners to learn how to build a web-based python Django project.

We will provide you with the complete source code and database for the python project so that you can easily install it on your machine and learn how to program in Python Django.

>[!NOTE]
> To start creating a **Payroll Management System Project in Python Django**, makes sure that you have PyCharm Professional IDE Installed in your computer.

## Admin Features of Payroll Management System Project in Django

* **Manage Employee**

For the employee, he admin can add, and edit, employee information.

* **Company Management**

For the company, the admin can see the list of company details. Admin can update the record of the company details.

* **Manage payroll**

For the payroll, the admin can view the list of payroll details. Admin can create a payroll to their employees.

* **Manage Voucher**

For the voucher, The admin can add, and edit, voucher information.

* **Login**

By default the admin need to login first to enable to access the system.

* **Manage logs**

For the logs, The admin can view all the logs.

* **Manage Company**

For the Company, The admin can add, edit, and delete company information.

Payroll Management System Project in Django Steps on How to Create a Project
Time needed: 5 minutes

Here are the steps on how to create a Payroll Management System Project in Django

1. **Open file**.

First, open “pycharm professional” after that click “file” and click “new project”.

![image](https://github.com/user-attachments/assets/f22e09a3-3478-457f-a150-3a006dcfae7b)

2. **Choose Django**.

Next, after click “new project”, choose “Django” and click.

![image](https://github.com/user-attachments/assets/a1a45535-689b-4b41-b152-a8e8f59d0d70)

3. **Select file location.**

Then, select a file location wherever you want.

4. **Create application name.**

After that, name your application.

5. **Click create**.

Lastly, finish creating project by clicking “create” button.

6. **Start Coding**.

Finally, we will now start adding functionality to our Django Framework by adding some functional codes.

##  Functionality and Codes of the Payroll Management System Project in Django

*  Create template for the login in form in Payroll Management System Project in Django.

In this section, we will learn on how create a templates for the login form. 

To start with, add the following code in your login.html under the folder of login/templates/login.

```
<html>
<header>
{% load static %}
<title>Login</title>
<link rel="stylesheet" type="text/css" href="{% static 'dashboard/css/bootstrap.min.css' %}">
<link href="https://fonts.googleapis.com/css?family=Raleway|Roboto+Condensed" rel="stylesheet">
<link rel="stylesheet" type="text/css" href="{% static 'myaccount/style.css' %}">
<script  href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js"></script>
<script
  src="https://code.jquery.com/jquery-3.2.1.min.js"
  integrity="sha256-hwg4gsxgFZhOsEEamdOYGBf13FyQuiTwlAQgxVSNgt4="
  crossorigin="anonymous"></script>
  
  </header>
  
<body>

<div class="container">
    <div class="row">
	    <div class="col-sm-6 col-sm-push-3 login-div">
		    {% if message %}
			<div class="alert alert-danger alert-dismissable "><a href="#" class="close" data-dismiss="alert" aria-label="close">&times;</a>{{ message }}</div>
			{% endif %}
			<div class='card'>
				<form method="post" class="">
				{% csrf_token %}
				{% for field in form %}
					<div class="form-group">
						{{field.label_tag}}
						{{ field }}
					</div>
				{% endfor %}
					<button type="submit" name="submit" class="btn btn-primary form-control"/>Login</button>
				</form>
			</div>
		</div>
	</div>
</div>
    
  
  
</body>
</html>
```

* **Create template for the employee table in Payroll Management System Project in Django**.

In this section, we will learn on how create a templates for the employee table. 

To start with, add the following code in your all_employees.html under the folder of employee/templates/employee.

```
{% extends "hrms/base.html" %}
{% load crispy_forms_tags %}
{% block content %}
	<div class="card">
		<div class="card-header">
			<h4 class="card-title">
			{{ head }}
			</h4>
		</div>
		<div class="card-body">

        <a href='{% url "employee-add" %}' class="btn btn-danger" >Add Employee</a>

        


				{% comment %} <form class="form-control" action="{% url 'search-employees' %}" method="post" style="margin-top: 10px">
						{% csrf_token %}
					<div class="row">
	            <div class="col example-grid-col">
								  <select class="form-control" name="company">
								  	<option value="">Choose company</option>
										{% for com in companies %}
											<option value="{{com.company_name}}">{{com.company_name}}</option>
										{% endfor %}
								  </select>
	            </div>
	            <div class="col example-grid-col">
	                <input type="text" class="form-control" name="employee_name" value="" placeholder="Search by employee">
	            </div>
	            
				<div class="col example-grid-col">
					<select class="form-control" name="payment_method">
						<option value="">Choose payment method</option>
							<option value="daily">daily</option>
							<option value="weekly">weekly</option>
							<option value="semi-monthly">semi-monthly</option>
							<option value="monthly">monthly</option>
					</select>
	            </div>

	            <div class="col example-grid-col">
	                <input type="text" name="start_date" value="" class="datepicker dateinput form-control"  id="id_date_hired" placeholder="Start Date">
	            </div>
	            <div class="col example-grid-col">
	                <input type="text" name="end_date" value="" class="datepicker dateinput form-control"  id="id_date_hired" placeholder="End Date">
	            </div>

	            <div class="col example-grid-col">
					<select class="form-control" name="search_by">
						<option value="">Choose Search By</option>
							<option value="date_hired">Date Hired</option>
							<option value="contract_expiration">Contract Expiration</option>
					</select>
	            </div>

	            <div class="col example-grid-col">
					<select class="form-control" name="order">
						<option value="">Choose Order</option>
							<option value="ascending">ascending</option>
							<option value="descending">descending</option>
					</select>
	            </div>

							<div class="col example-grid-col">
	                <button type="submit" class="btn btn-info" name="button">Search</button>
	            </div>
	        </div>
				</form> {% endcomment %}
				<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/4.1.3/css/bootstrap.css">
			<link rel="stylesheet" href="https://cdn.datatables.net/1.10.19/css/dataTables.bootstrap4.min.css">

			<div class="table-responsive" style="margin-top: 10px">
				<table id="example" class="table table-striped table-bordered" style="width:100%">

					<thead class="bg-success text-white">
						<tr>
							<th scope="col" class="text-center">#</th>
							<th scope="col" class="text-center">Emp Id</th>
							<th scope="col" class="text-center">Employee Name</th>
							<th scope="col" class="text-center">SSS</th>
							<th scope="col" class="text-center">Phone</th>
							<th scope="col" class="text-center">Position</th>
							<th scope="col" class="text-center">Company</th>
							<th scope="col" class="text-center">Date Hired</th>
							<th scope="col" class="text-center">Contact Expiration</th>
							<th scope="col" class="text-center">Remarks</th>
							<th scope="col" class="text-center">Action</th>
						</tr>
					</thead>
					<tbody>
						{% for emp in employees %}
							<tr>
								<th class="align-middle text-center" scope="row">{{ forloop.counter}}</th>
								<td class="align-middle text-center">{{ emp.emp_id }} </td>
								<td class="align-middle text-center">{{ emp }} </td>
								<td class="align-middle text-center">{{ emp.get_hiring_details.sss }}</td>
								<td class="align-middle text-center">{{ emp.phone }}</td>
								<td class="align-middle text-center">{{ emp.get_hiring_details.position }}</td>
								<td class="align-middle text-center">{{ emp.company }}</td>
								<td class="align-middle text-center">{{ emp.date_hired }}</td>
								<td class="align-middle text-center">{{ emp.contract_expiration }}</td>
								<td class="align-middle text-center">{{ emp.remarks }}</td>
								<td class="align-middle text-center">
									<a href="{% url 'employee-update' emp.id %}" class="btn btn-danger">Update</a>
								</td>
							</tr>
						{% endfor %}
					</tbody>
				</table>
			</div>
		</div>
	</div>
{% endblock content %}
```

### 📌Here's the full documentation for the [Payroll Management System Project in Django](https://itsourcecode.com/free-projects/python-projects/payroll-management-system-project-in-django-with-source-code/)
