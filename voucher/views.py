from django.shortcuts import render, redirect, get_object_or_404
from django.http import HttpResponse, HttpResponseRedirect
from django.contrib import messages
from django.db.models import Q
from itertools import chain
from django.contrib import messages
from .models import Voucher, Voucher_particulars
from logs.models import Logs
from .forms import VoucherForm, ParticularsForm
from django.contrib.auth.decorators import login_required
from datetime import datetime, timedelta
import arrow

import csv
import xlwt


@login_required
def download_voucher(request, voucher_id):
    voucher = Voucher.objects.filter(pk=voucher_id).first()
    particulars = Voucher_particulars.objects.filter(voucher=voucher_id)
    response = HttpResponse(content_type='text/ms-excel')
    file_name = f"{voucher.voucher_no} - Cash Voucher.xls"
    response['Content-Disposition'] = f'attachment; filename="{file_name}"'
    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('Cash Voucher')
    style_bold = 'align: wrap on, vert centre, horiz center; font: bold on'
    style_square_center = 'align: wrap on, vert center, horiz center; font: bold on, color black; borders: top_color black, bottom_color black, right_color black, left_color black,\
	      left thin, right thin, top thin, bottom thin; pattern: pattern solid, fore_color white;'
    style_top_left_right = 'font: height 175; align: wrap on, vert center, horiz center; font: bold on, color black; borders: top_color black, right_color black, left_color black,\
	      left thin, right thin, top thin; pattern: pattern solid, fore_color white;'
    style_left = 'font: height 150; align: wrap on, vert center, horiz center; font: bold on, color black; borders: left_color black,\
	      left thin; pattern: pattern solid, fore_color white;'
    style_right = 'font: height 150; align: wrap on, vert center, horiz center; font: color black; borders: right_color black,\
	      right thin; pattern: pattern solid, fore_color white;'
    style_left_right = 'font: height 175; align: wrap on, vert center, horiz center; font: bold on, color black; borders: left_color black, right_color black,\
	      left thin, right thin; pattern: pattern solid, fore_color white;'
    style_bottom_right = 'font: height 175; align: wrap on, vert center, horiz center; font: bold on, color black; borders: bottom_color black, right_color black,\
	      right thin, bottom thin; pattern: pattern solid, fore_color white;'
    style_bottom = 'font: height 175; align: wrap on, vert center, horiz center; font: bold on, color black; borders: bottom_color black,\
	      bottom thin; pattern: pattern solid, fore_color white;'
    style_top = 'font: height 175; align: wrap on, vert center, horiz center; font: bold on, color black; borders: top_color black,\
	      top thin; pattern: pattern solid, fore_color white;'

    style_normal = 'font: height 175; align: wrap on, vert centre, horiz center;'
    ws.write_merge(1, 1, 0, 0, 'RC No.',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(1, 1, 1, 2, f'{voucher.rc_no	}',
                   xlwt.Style.easyxf(style_bottom))
    style_for_title = 'font: height 300;align: wrap on, vert center, horiz center; font: bold on, color black; borders: top_color black, bottom_color black, right_color black, left_color black,\
 bottom thin; pattern: pattern solid, fore_color white;'

    ws.write_merge(0, 1, 4, 6, 'CASH VOUCHER',
                   xlwt.Style.easyxf(style_for_title))

    ws.write_merge(1, 1, 8, 8, 'No.',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(1, 1, 9, 10, f'{voucher.voucher_no}',
                   xlwt.Style.easyxf(style_bottom))

    ws.write_merge(2, 2, 0, 0, 'Date',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(2, 2, 1, 2, f'{voucher.date_created}',
                   xlwt.Style.easyxf(style_bottom))

    ws.write_merge(2, 2, 8, 8, 'Date',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(2, 2, 9, 10, f'{voucher.voucher_created_date}',
                   xlwt.Style.easyxf(style_bottom))

    ws.write_merge(3, 3, 0, 0, 'Place',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(3, 3, 1, 2, f'{voucher.place}',
                   xlwt.Style.easyxf(style_bottom))

    ws.write_merge(4, 4, 0, 0, 'Paid to',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(4, 4, 1, 6, f'{voucher.paid_to}',
                   xlwt.Style.easyxf(style_bottom))

    ws.write_merge(5, 5, 0, 0, 'Address',
                   xlwt.Style.easyxf(style_normal))
    ws.write_merge(5, 5, 1, 6, f'{voucher.address}',
                   xlwt.Style.easyxf(style_bottom))

    ws.write_merge(7, 7, 0, 7, 'Particulars',
                   xlwt.Style.easyxf(style_square_center))
    ws.write_merge(7, 7, 8, 10, f'Amount',
                   xlwt.Style.easyxf(style_square_center))

    rowrow = 7
    total_amount = 0
    for part in particulars:
        rowrow = rowrow + 1
        total_amount = total_amount + part.amount
        ws.write_merge(rowrow, rowrow, 0, 7, f'{part.particular_name}',
                       xlwt.Style.easyxf(style_normal))
        ws.write_merge(rowrow, rowrow, 8, 9, f'{part.amount}',
                       xlwt.Style.easyxf(style_left_right))

    remaining_row = 10 - rowrow

    for i in range(1, remaining_row):
        rowrow = rowrow + 1
        ws.write_merge(rowrow, rowrow, 8, 9, f'',
                       xlwt.Style.easyxf(style_left_right))

    rowrow = rowrow + 1
    ws.write_merge(rowrow, rowrow, 6, 7, f'TOTAL PHP',
                   xlwt.Style.easyxf(style_normal))

    ws.write_merge(rowrow, rowrow, 8, 9, f'{total_amount}',
                   xlwt.Style.easyxf(style_left_right))
    rowrow = rowrow + 1
    ws.write_merge(rowrow, rowrow, 0, 10, f'',
                   xlwt.Style.easyxf(style_top))
    rowrow = rowrow + 1
    ws.write_merge(rowrow, rowrow, 0, 10, f'RECEIVED from ____________________________the amount of',
                   xlwt.Style.easyxf(style_normal))
    rowrow = rowrow + 1
    ws.write_merge(rowrow, rowrow, 0, 10, f'PESOS ________________________________(PHP________)',
                   xlwt.Style.easyxf(style_normal))
    rowrow = rowrow + 1
    ws.write_merge(rowrow, rowrow, 0, 3, f'',
                   xlwt.Style.easyxf(style_bottom))
    ws.write_merge(rowrow, rowrow, 5, 8, f'in full payment of amount described above.',
                   xlwt.Style.easyxf(style_bottom))
    rowrow = rowrow + 1
    ws.write_merge(rowrow, rowrow, 0, 3, f'Approved',
                   xlwt.Style.easyxf(style_normal))
    rowrow = rowrow + 1
    ws.write_merge(rowrow, rowrow, 7, 10, f'By:___________________________',
                   xlwt.Style.easyxf(style_normal))

    action = f"Voucher downloaded with  RC# {voucher.rc_no}"
    Logs.objects.create(
        action=action, action_by=request.user, action_date=datetime.now())
    wb.save(response)
    return response


@login_required
def delete_particulars(request, particular_id):
    particular = get_object_or_404(Voucher_particulars, pk=particular_id)
    voucher_id = particular.voucher.id
    action = f"Particular was successfully deleted in voucher RC# {particular.voucher.rc_no} - . (particular_name - {particular.particular_name}, amount - {particular.amount})"
    Logs.objects.create(
        action=action, action_by=request.user, action_date=datetime.now())
    particular.delete()
    messages.success(request, "Particular was successfully deleted")
    return redirect("/vouchers/encode-particulars/" + str(voucher_id))


@login_required
def encode_particulars(request, voucher_id):
    voucher = get_object_or_404(Voucher, pk=voucher_id)

    if request.method == "POST":
        form = ParticularsForm(request.POST)
        if form.is_valid():
            particulars_form = form.save(commit=False)
            particulars_form.voucher_id = voucher_id
            # particulars_form.action_by = request.user
            particulars_form.save()
            action = f"Particular was successfully encoded in voucher RC# {voucher.rc_no} - . (particular_name - {form.cleaned_data['particular_name']}, amount - {form.cleaned_data['amount']})"
            Logs.objects.create(
                action=action, action_by=request.user, action_date=datetime.now())
            messages.success(request, "Particular was successfully added")
            return redirect(f"/vouchers/encode-particulars/{voucher_id}")
        else:
            print(f" errors - {form.errors}")
    particulars = Voucher_particulars.objects.filter(voucher=voucher_id)
    form = ParticularsForm()

    context = {
        "head": f" Particulars for voucher no: {voucher.voucher_no}",
        "particulars": particulars,
        "form": form,
        "voucher_id": voucher_id,
        "voucher": voucher
    }
    return render(request, "voucher/particulars.html", context)


@login_required
def voucher_update(request, pk):
    voucher = get_object_or_404(Voucher, pk=pk)
    if request.method == 'POST':
        form = VoucherForm(request.POST or None, instance=voucher)
        if form.is_valid():
            form.save()
            action = f"Voucher was successfully updated. (rc no - {form.cleaned_data['rc_no']}, date_created - {form.cleaned_data['date_created']}, place - {form.cleaned_data['place']}, voucher_no - {form.cleaned_data['voucher_no']}, voucher_created_date - {form.cleaned_data['voucher_created_date']}, paid_to - {form.cleaned_data['paid_to']}, address - {form.cleaned_data['address']})"
            Logs.objects.create(
                action=action, action_by=request.user, action_date=datetime.now())
            messages.success(request, 'Voucher was successfully updated.')
            return redirect('voucher-update', pk=pk)
    else:
        form = VoucherForm(instance=voucher)

    context = {
        'head': 'Update Voucher',
        'form': form,
        'for_update': 1,
        'pk': pk,
                'voucher': voucher
    }
    return render(request, 'voucher/voucher_add.html', context)


@login_required
def voucher(request):
    vouchers = Voucher.objects.all()
    context = {
        "head": "Vouchers",
        "vouchers": vouchers
    }
    # return HttpResponse(vouchers)
    return render(request, 'voucher/vouchers.html', context)


@login_required
def voucher_add(request):
    if request.method == 'POST':
        form = VoucherForm(request.POST)
        if form.is_valid():
            form.save()
            action = f"Voucher was successfully added. (rc no - {form.cleaned_data['rc_no']}, date_created - {form.cleaned_data['date_created']}, place - {form.cleaned_data['place']}, voucher_no - {form.cleaned_data['voucher_no']}, voucher_created_date - {form.cleaned_data['voucher_created_date']}, paid_to - {form.cleaned_data['paid_to']}, address - {form.cleaned_data['address']})"
            Logs.objects.create(
                action=action, action_by=request.user, action_date=datetime.now())
            messages.success(
                request, f'Voucher has been successfully created.')
            return redirect('voucher')
    else:
        form = VoucherForm()
    context = {
        'head': 'Add Voucher',
        'form': form
    }
    return render(request, 'voucher/voucher_add.html', context)
