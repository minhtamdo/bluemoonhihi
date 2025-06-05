"""
URL configuration for bluemoon project.

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/5.0/topics/http/urls/
Examples:
Function views
    1. Add an import:  from my_app import views
    2. Add a URL to urlpatterns:  path('', views.home, name='home')
Class-based views
    1. Add an import:  from other_app.views import Home
    2. Add a URL to urlpatterns:  path('', Home.as_view(), name='home')
Including another URLconf
    1. Import the include() function: from django.urls import include, path
    2. Add a URL to urlpatterns:  path('blog/', include('blog.urls'))
"""
from django.contrib import admin
from django.urls import path
from django.shortcuts import render
from core.models import *  # sửa lại path import theo app của bạn
from django.utils import timezone
from django.db.models import Sum, Count, Q, OuterRef, Subquery, Exists, F
from datetime import datetime
from django.template.loader import render_to_string
from django.http import JsonResponse, HttpResponse
from core.forms import FeeForm
from datetime import date, time
import calendar
from openpyxl import Workbook
from django.http import HttpResponse
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_POST
from django.contrib.auth.decorators import login_required
from datetime import datetime
from django.utils.timezone import localtime
from django.shortcuts import redirect, get_object_or_404
from django.db import IntegrityError
from django.core.exceptions import ValidationError

import traceback
import json
from django.db import connection
from django.shortcuts import redirect


@csrf_exempt
def homeowner_view(request):
    if not request.session.get('user_id') or request.session.get('role') != 'chu_ho':
        return redirect('/login')

    user_id = request.session['user_id']
    user = User.objects.get(id=user_id)
    household = Household.objects.filter(head=user).first()

    # Phí chung hoặc phí riêng áp dụng cho hộ đó
    applicable_fees = Fee.objects.filter(
        is_common=True
    ) | Fee.objects.filter(households=household)

    fee_status_list = []
    for fee in applicable_fees.distinct():
        payment = Payment.objects.filter(fee=fee, household=household).first()
        status = payment.status if payment else 'unpaid'
        fee_status_list.append({
            'title': fee.title,
            'amount': f"${fee.amount:,.0f}",
            'status': status
        })

    return render(request, 'homeowner.html', {
        'fullname': user.fullname,
        'apartment_number': household.household_number,
        'fee_status_list': fee_status_list
    })

@csrf_exempt
def ketoan_view(request):
    if not request.session.get('user_id') or request.session.get('role') != 'thu_ky':
        return redirect('/login')
    
    if request.method == 'GET' and request.headers.get('x-requested-with') == 'XMLHttpRequest' and 'month' in request.GET:
        month = request.GET.get("month")
        try:
            year, month = map(int, month.split("-"))
            start_date = datetime.date(year, month, 1)
            last_day = calendar.monthrange(year, month)[1]
            end_date = datetime.date(year, month, last_day)
        except:
            return HttpResponse("Tháng không hợp lệ", status=400)

    # Lấy các khoản phí đến hạn trong tháng
        fees = Fee.objects.filter(due_date__range=(start_date, end_date)).order_by('due_date')

    # Tạo Excel
        wb = Workbook()
        ws = wb.active
        ws.title = "Báo cáo phí"

        ws.append(["Căn hộ", "Loại phí", "Số tiền", "Hạn đóng", "Trạng thái", "Ngày thanh toán"])

        for fee in fees:
            payments = fee.payments.all()
            for payment in payments:
                ws.append([
                    payment.household.household_number,
                    fee.title,
                    float(fee.amount),
                    fee.due_date.strftime("%d/%m/%Y"),
                    payment.get_status_display(),
                    payment.paid_at.strftime("%d/%m/%Y %H:%M") if payment.paid_at else ''
                ])

        response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        filename = f"bao_cao_phi_{year}_{month:02}.xlsx"
        response['Content-Disposition'] = f'attachment; filename=\"{filename}\"'
        wb.save(response)
        return response

    # Tổng số chủ hộ
    total_residents = User.objects.filter(role='chu_ho').count()
    now = timezone.localtime(timezone.now())
    today = now.date()

    # Xác định mốc thời gian
    start_of_this_month = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)

    if now.month == 12:
        start_of_next_month = start_of_this_month.replace(year=now.year + 1, month=1)
    else:
        start_of_next_month = start_of_this_month.replace(month=now.month + 1)

    if now.month == 1:
        start_of_last_month = start_of_this_month.replace(year=now.year - 1, month=12)
    else:
        start_of_last_month = start_of_this_month.replace(month=now.month - 1)

    # Tổng doanh thu tháng này
    total_revenue = Payment.objects.filter(
        paid_at__gte=start_of_this_month,
        paid_at__lt=start_of_next_month,
        status='paid'
    ).aggregate(total=Sum('fee__amount'))['total'] or 0

    # Doanh thu tháng trước
    last_month_revenue = Payment.objects.filter(
        paid_at__gte=start_of_last_month,
        paid_at__lt=start_of_this_month,
        status='paid'
    ).aggregate(total=Sum('fee__amount'))['total'] or 0

    # Tính phần trăm thay đổi
    if last_month_revenue == 0:
        change_percent = None
    else:
        change_percent = ((total_revenue - last_month_revenue) / last_month_revenue) * 100

    # Đã thanh toán trong tháng này
    paid_count = Payment.objects.filter(
        paid_at__gte=start_of_this_month,
        paid_at__lt=start_of_next_month,
        status='paid'
    ).count()

    # Chưa thanh toán trong tháng này
    unpaid_count = Payment.objects.filter(
        paid_at__gte=start_of_this_month,
        paid_at__lt=start_of_next_month,
        status='unpaid'
    ).count()

    # Tổng phí cần thu trong tháng
    total_due = Fee.objects.filter(
        due_date__gte=start_of_this_month,
        due_date__lt=start_of_next_month
    ).aggregate(total=Sum('amount'))['total'] or 0

    # Tỷ lệ thu
    collection_rate = (total_revenue / total_due) * 100 if total_due > 0 else None

    # Lấy danh sách nhân khẩu
    members = HouseholdMember.objects.all()

    # Lấy số hoá đơn quá hạn
    overdue_bills = Fee.objects.filter(
    due_date__lt=today
    ).exclude(
    payments__status='paid'
    ).distinct().count()

    # Tổng nợ: các hóa đơn quá hạn chưa thanh toán
    paid_subquery = Payment.objects.filter(
    fee=OuterRef('pk'),
    status='paid'
    )

    total_debt = Fee.objects.filter(
        due_date__lt=today
    ).annotate(
        has_paid=Exists(paid_subquery)
    ).filter(
        has_paid=False
    ).aggregate(
        total=Sum('amount')
    )['total'] or 0

    households_need_contact = Household.objects.filter(
    payments__fee__due_date__lt=today,
    payments__status='unpaid'
    ).distinct()

    need_contact_count = households_need_contact.count()
    fees = Fee.objects.all()
    payments = Payment.objects.select_related('fee', 'household').all()

    households = Household.objects.all()
    form = FeeForm()
    if request.method == 'POST' and request.headers.get('x-requested-with') == 'XMLHttpRequest':
        form = FeeForm(request.POST)
        if form.is_valid():
            user_id = request.session.get('user_id')
            if not user_id:
                return JsonResponse({'success': False, 'message': 'Người dùng chưa đăng nhập.'}, status=403)

            try:
                user = User.objects.get(id=user_id)
            except User.DoesNotExist:
                return JsonResponse({'success': False, 'message': 'Không tìm thấy người dùng.'}, status=404)

            fee = form.save(commit=False)
            fee.created_by = user
            fee.created_at = datetime.now()

            # Xử lý phí chung hay riêng
            is_common = request.POST.get('is_common', 'true') == 'true'
            fee.is_common = is_common
            fee.save()

            if not is_common:
                household_ids = request.POST.getlist('households') # Nếu cần import
                households = Household.objects.filter(id__in=household_ids)
                fee.households.set(households)  # thiết lập quan hệ N-N

            return JsonResponse({
                'success': True,
                'id': str(fee.id),
                'title': fee.title,
                'amount': float(fee.amount),
                'type': fee.type,
                'due_date': fee.due_date.strftime('%Y-%m-%d'),
                'description': fee.description or '',
                'is_common': fee.is_common
            })
        return JsonResponse({'success': False, 'errors': form.errors}, status=400)

    # Nếu là GET, render trang kế toán
    if not request.session.get('user_id') or request.session.get('role') != 'thu_ky':
        return redirect('/login')
    
    if request.method == 'GET' and request.headers.get('x-requested-with') == 'XMLHttpRequest' and request.GET.get('debt') == '1':

        today = now.date()

        unpaid_payments = Payment.objects.filter(
            status='unpaid',
            fee__due_date__lt=today
        ).select_related('fee', 'household__head')

        wb1 = Workbook()
        ws1 = wb1.active
        ws1.title = "Danh sách nợ"

        ws1.append(["Căn hộ", "Chủ hộ", "Loại phí", "Số tiền", "Hạn đóng"])

        for payment in unpaid_payments:
            household = payment.household
            head_name = household.head.fullname if household and household.head else "Không rõ"
            fee = payment.fee
            ws1.append([
                household.household_number,
                head_name,
                fee.title,
                float(fee.amount),
                fee.due_date.strftime("%d/%m/%Y")
            ])

        response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        response['Content-Disposition'] = 'attachment; filename="danh_sach_no.xlsx"'
        wb1.save(response)
        return response
    user_id = request.session.get('user_id')
    user = User.objects.get(id=user_id)


    return render(request, 'ketoan.html', {
        'total_residents': total_residents,
        'total_revenue': total_revenue,
        'change_percent': change_percent,
        'paid_count': paid_count,
        'unpaid_count': unpaid_count,
        'collection_rate': collection_rate,
        'members': members,
        'overdue_bills': overdue_bills,
        'total_debt' : total_debt,
        'need_contact_count': need_contact_count,
        'fees': fees,
        'payments': payments,
        'fullname': user.fullname,
        'households': households,
        'form': form,
    })

def update_payment(request):
    if request.method == 'POST':
        try:
            data = json.loads(request.body)
            payment_id = data.get('payment_id')
            paid_at = data.get('paid_at')
            status = data.get('status')
            method = data.get('method')

            payment = Payment.objects.get(id=payment_id)
            payment.paid_at = paid_at
            payment.status = status
            payment.method = method
            payment.save()

            return JsonResponse({'success': True})
        except Payment.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'Payment không tồn tại'})
        except Exception as e:
            return JsonResponse({'success': False, 'error': str(e)})

    return JsonResponse({'success': False, 'error': 'Phương thức không hợp lệ'})

def revenue_pie_data(request):
    now = timezone.localtime(timezone.now())
    today = now.date()
    current_month = today.month
    current_year = today.year

    data = (
        Payment.objects
        .filter(
            status='paid',
            paid_at__year=current_year,
            paid_at__month=current_month
        )
        .values(title=F('fee__title'))
        .annotate(total=Sum('fee__amount'))
    )

    total_amount = sum(item['total'] for item in data)
    response_data = [
        {
            'title': item['title'],
            'value': round(item['total'] / total_amount * 100, 1) if total_amount > 0 else 0
        } for item in data
    ]

    return JsonResponse(response_data, safe=False)

@csrf_exempt
def login_view(request):
    
    if request.method == 'GET':
        return render(request, 'dangnhap.html')

    if request.method == 'POST':
        try:
            data = json.loads(request.body)
            username = data.get('username')
            password = data.get('password')
            role_input = data.get('role')

            role_map = {
                'resident': 'chu_ho',
                'leader': ['to_truong', 'to_pho'],
                'accountant': 'thu_ky'
            }
            allowed_roles = role_map.get(role_input)
            if isinstance(allowed_roles, str):
                allowed_roles = [allowed_roles]

            with connection.cursor() as cursor:
                cursor.execute("""
                    SELECT user_id, username, fullname, role
                    FROM users
                    WHERE username = %s
                    AND role IN %s
                    AND password_hash = crypt(%s, password_hash)
                    AND status = 'active'
                """, [username, tuple(allowed_roles), password])
                row = cursor.fetchone()

            if row:
                # ✔️ Lưu trạng thái đăng nhập
                request.session['user_id'] = str(row[0])
                request.session['username'] = row[1]
                request.session['fullname'] = row[2]
                request.session['role'] = row[3]

                return JsonResponse({
                    'success': True,
                    'redirect_url': get_redirect_url(row[3]),
                    'user': {'fullname': row[2], 'username': row[1], 'role': row[3]}
                })
            else:
                return JsonResponse({'success': False, 'message': 'Sai tài khoản, mật khẩu hoặc vai trò không hợp lệ.'})
        except Exception as e:
            return JsonResponse({'success': False, 'message': str(e)})
    return JsonResponse({'success': False, 'message': 'Chỉ hỗ trợ POST'})

def logout_view(request):
    request.session.flush()
    return redirect('/login')

@csrf_exempt
def leader_view(request):
    if not request.session.get('user_id') or (request.session.get('role') != 'to_truong' and request.session.get('role') != 'to_pho'):
        return redirect('/login')
    total_households = Household.objects.count()
    total_residents = HouseholdMember.objects.count()
    total_pending = ResidencyRequest.objects.filter(status='pending').count()
    households = Household.objects.all()
    residents = HouseholdMember.objects.select_related('household').all()
    residency_requests = ResidencyRequest.objects.select_related('user').order_by('-created_at')
    total = (
    HouseholdMember.objects.count()
    + ResidencyRequest.objects.filter(status='approved', request_type='temporary_residence').count()
    - ResidencyRequest.objects.filter(status='approved', request_type='temporary_absence').count())
    context = {
        'total_households': total_households,
        'total_residents': total_residents,
        'total_pending': total_pending,
        'households': households,
        'residents': residents,
        'residency_requests': residency_requests,
        'total': total
    }

    return render(request, 'leader.html', context)

@csrf_exempt
def approve_request(request, request_id):
    if request.method == 'POST':
        try:
            data = json.loads(request.body)
            action = data.get('action')

            if action == 'approve':
                req = ResidencyRequest.objects.get(id=request_id)
                req.status = 'approved'
                req.save()
                return JsonResponse({'success': True})
            else:
                return JsonResponse({'success': False, 'message': 'Hành động không hợp lệ'})
        except Exception as e:
            return JsonResponse({'success': False, 'message': str(e)})

    return JsonResponse({'success': False, 'message': 'Phương thức không hợp lệ'})

@csrf_exempt
def reject_request(request, request_id):
    if request.method == 'POST':
        try:
            data = json.loads(request.body)
            action = data.get('action')

            if action == 'reject':
                req = ResidencyRequest.objects.get(id=request_id)
                req.status = 'rejected'
                req.save()
                return JsonResponse({'success': True})
            else:
                return JsonResponse({'success': False, 'message': 'Hành động không hợp lệ'})
        except Exception as e:
            return JsonResponse({'success': False, 'message': str(e)})

    return JsonResponse({'success': False, 'message': 'Phương thức không hợp lệ'})

@csrf_exempt
def editHousehold(request, household_id):
    if request.method == 'POST':
        try:
            household = get_object_or_404(Household, id=household_id)

            household.household_number = request.POST.get('household_number', '').strip()
            household.head_name = request.POST.get('head_name', '').strip()
            household.address = request.POST.get('address', '').strip()

            household.save()
            return JsonResponse({'success': True})
        except Exception as e:
            return JsonResponse({'success': False, 'message': str(e)}, status=500)
    
    return JsonResponse({'success': False, 'message': 'Phương thức không được hỗ trợ'}, status=405)

@csrf_exempt
def deleteHousehold(request, household_id):
    if request.method == 'POST':
        try:
            household = get_object_or_404(Household, id=household_id)
            household.delete()
            return JsonResponse({'success': True})
        except Exception as e:
            return JsonResponse({'success': False, 'message': str(e)}, status=500)

    return JsonResponse({'success': False, 'message': 'Phương thức không được hỗ trợ'}, status=405)

@csrf_exempt
def editResident(request, member_id):
    if request.method == 'POST':
        try:
            member = get_object_or_404(HouseholdMember, id=member_id)

            member.full_name = request.POST.get('full_name', '').strip()
            member.gender = request.POST.get('gender', '').strip()
            member.dob = request.POST.get('dob', '').strip()
            member.other_name = request.POST.get('other_name', '').strip()
            member.household_number = request.POST.get('household_number', '').strip()
            member.relationship = request.POST.get('relationship', '').strip()
            member.place_of_birth = request.POST.get('place_of_birth', '').strip()
            member.native_place = request.POST.get('native_place', '').strip()
            member.ethnic = request.POST.get('ethnic', '').strip()
            member.occupation = request.POST.get('occupation', '').strip()
            member.cccd = request.POST.get('cccd', '').strip()
            member.place_of_work = request.POST.get('place_of_work', '').strip()
            member.issue_date = request.POST.get('issue_date', '').strip()
            member.issued_by = request.POST.get('issued_by', '').strip()
            member.note = request.POST.get('note', '').strip()
            member.is_temporary = request.POST.get('is_temporary', 'false').lower() == 'true'
            member.joined_at = request.POST.get('joined_at', '').strip()

            member.save()
            return JsonResponse({'success': True})
        except Exception as e:
            return JsonResponse({'success': False, 'message': str(e)}, status=500)

    return JsonResponse({'success': False, 'message': 'Phương thức không được hỗ trợ'}, status=405)

@csrf_exempt
def deleteResident(request, member_id):
    if request.method == 'POST':
        try:
            member = get_object_or_404(HouseholdMember, id=member_id)
            member.delete()
            return JsonResponse({'success': True})
        except Exception as e:
            return JsonResponse({'success': False, 'message': str(e)}, status=500)

    return JsonResponse({'success': False, 'message': 'Phương thức không được hỗ trợ'}, status=405)

def get_users(request):
    users = User.objects.all()
    data = [{"id": str(u.id), "fullname": u.fullname, "username": u.username} for u in users]
    return JsonResponse(data, safe=False)

@csrf_exempt
def add_household(request):
    if request.method == "POST":
        try:
            data = json.loads(request.body)

            head_id = data.get("head_id")
            head = User.objects.get(id=head_id)

            household = Household.objects.create(
                household_number=data.get("household_number"),
                head=head,
                household_size=int(data.get("household_size")),
                address=data.get("address"),
                created_at=timezone.now(),
                updated_at=timezone.now()
            )

            return JsonResponse({
                "id": str(household.id),
                "household_number": household.household_number,
                "head_name": head.fullname,
                "household_size": household.household_size,
                "address": household.address,
                "created_at": household.created_at.strftime("%d/%m/%Y"),
                "updated_at": household.updated_at.strftime("%d/%m/%Y"),
            })

        except User.DoesNotExist:
            return JsonResponse({"error": "Không tìm thấy chủ hộ"}, status=400)
        except Exception as e:
            return JsonResponse({"error": str(e)}, status=500)

    return JsonResponse({"error": "Phương thức không hợp lệ"}, status=405)

@csrf_exempt
def add_resident(request):
    if request.method == 'POST':
        data = request.POST

        try:
            household_number = data.get('household_number')
            household = Household.objects.get(household_number=household_number)

            def parse_date(date_str):
                return datetime.strptime(date_str, '%Y-%m-%d').date() if date_str else None

            full_name = data.get('full_name')
            gender = data.get('gender')
            dob = parse_date(data.get('dob'))
            relationship = data.get('relationship')
            joined_at = timezone.now()

            # ✅ Kiểm tra trường bắt buộc
            if not all([full_name, gender, dob, relationship, joined_at]):
                return JsonResponse({'status': 'error', 'message': 'Thiếu thông tin bắt buộc'}, status=400)

            # ✅ Tạo member
            member = HouseholdMember.objects.create(
                household=household,
                full_name=data.get('full_name'),
                other_name=data.get('other_name'),
                gender=data.get('gender'),
                dob=parse_date(data.get('dob')),
                place_of_birth=data.get('place_of_birth'),
                native_place=data.get('native_place'),
                ethnic_group=data.get('ethnic_group'),
                occupation=data.get('occupation'),
                place_of_work=data.get('place_of_work'),
                cccd=data.get('cccd'),
                issue_date=parse_date(data.get('issue_date')),
                issued_by=data.get('issued_by'),
                relationship=data.get('relationship'),
                is_temporary=False,  # bạn có thể chỉnh nếu muốn checkbox
                note=data.get('note'),
                joined_at=parse_date(data.get('joined_at')),
            )

            return JsonResponse({'status': 'success', 'id': str(member.id)})

        except Household.DoesNotExist:
            return JsonResponse({'status': 'error', 'message': 'Không tìm thấy hộ khẩu'}, status=404)
        except Exception as e:
            traceback.print_exc()
            return JsonResponse({'status': 'error', 'message': str(e)}, status=500)

def get_household_info(request, household_id):
    try:
        household = Household.objects.get(id=household_id)
    except Household.DoesNotExist:
        return JsonResponse({'error': 'Household not found'}, status=404)

    members = HouseholdMember.objects.filter(household=household)
    head_user = household.head

    data = {
        'owner': {
            'id': str(head_user.id) if head_user else None,
            'name': head_user.fullname if head_user else '(Không xác định)'
        },
        'members': [
            {
                'id': str(member.id),
                'name': member.full_name,
                'is_owner': head_user and member.full_name == head_user.fullname  # So sánh theo tên
            }
            for member in members
        ]
    }
    return JsonResponse(data)

@csrf_exempt
def delete_member(request, member_id):
    if request.method != 'POST':
        return JsonResponse({'error': 'Invalid request'}, status=400)

    try:
        member = HouseholdMember.objects.get(id=member_id)
        household = member.household
        is_owner = household.head and member.full_name == household.head.fullname  # Hoặc dùng CCCD nếu muốn chính xác hơn
    except HouseholdMember.DoesNotExist:
        return JsonResponse({'error': 'Member not found'}, status=404)

    if is_owner:
        household.delete()
    else:
        member.household = None
        member.save()

    return JsonResponse({'success': True})

def get_recent_activities(request):
    activities = []
    STATUS_LABELS = {
    'approved': 'Đồng ý',
    'rejected': 'Từ chối',
    'pending': 'Đang xử lý'
    }
    # 1. Lấy 10 hoạt động mới nhất từ Household
    households = Household.objects.all()
    for h in households:
        updated_at = h.updated_at
        if timezone.is_naive(updated_at):
            updated_at = timezone.make_aware(updated_at)
        activities.append({
            'time': updated_at.strftime("%d/%m/%Y %H:%M") if updated_at else '',
            'activity': f"Cập nhật hộ khẩu - Số nhà {h.household_number}",
            'status': 'Thành công',
            'updated_at': updated_at,
        })

    # 2. Lấy 10 hoạt động mới nhất từ HouseholdMember
    members = HouseholdMember.objects.select_related('household').all()
    for m in members:
        if m.joined_at:
            dt_naive = datetime.combine(m.joined_at, time.min)
            joined_datetime = timezone.make_aware(dt_naive)
        else:
            joined_datetime = timezone.make_aware(datetime(1900, 1, 1))
        activities.append({
            'time': m.joined_at.strftime("%d/%m/%Y") if m.joined_at else '',
            'activity': f"Cập nhật nhân khẩu - Tòa {m.household.household_number if m.household else '(Không xác định)'}",
            'status': 'Thành công',
            'updated_at': joined_datetime,
        })

    # 3. Lấy 10 hoạt động mới nhất từ ResidencyRequest
    requests = ResidencyRequest.objects.all()
    for r in requests:
        if timezone.is_naive(updated_at):
            updated_at = timezone.make_aware(updated_at)
    
    # Lấy tên người yêu cầu, nếu User có get_full_name, nếu không thì lấy username
    requester_name = r.user.fullname if r.user else "Người dùng"

    if r.request_type == 'temporary_absence':
        activity_desc = f"Yêu cầu tạm vắng của {requester_name}"
    elif r.request_type == 'temporary_residence':
        activity_desc = f"Yêu cầu tạm trú của {requester_name}"
    else:
        activity_desc = f"Yêu cầu khác: {r.request_type}"

    status_label = STATUS_LABELS.get(r.status, r.status)
    activities.append({
        'time': updated_at.strftime("%d/%m/%Y %H:%M") if updated_at else '',
        'activity': activity_desc,
        'status': status_label,
        'updated_at': updated_at,
    })


    # Sort toàn bộ theo updated_at datetime (giá trị aware), giảm dần (mới nhất lên đầu)
    activities.sort(key=lambda x: x['updated_at'], reverse=True)

    # Lấy 10 hoạt động gần đây nhất
    recent_activities = activities[:5]

    # Trả về JSON, bỏ trường updated_at không cần thiết cho frontend
    for act in recent_activities:
        act.pop('updated_at', None)

    return JsonResponse({'activities': recent_activities})

def get_redirect_url(role):
    if role == 'chu_ho':
        return '/cudan'
    elif role in ['to_truong', 'to_pho']:
        return '/leader'
    elif role == 'thu_ky':
        return '/ketoan'
    return '/login'

urlpatterns = [
    path('login/', login_view, name='home'),
    path('logout/', logout_view, name='logout'),
    path('admin/', admin.site.urls),
    path('ketoan/', ketoan_view, name='ketoan'),
    path('cudan/', homeowner_view, name='cudan'),
    path('leader/', leader_view, name='leader'),
    path('update_payment/', update_payment, name='update_payment'),
    path('revenue_pie_data/', revenue_pie_data, name='revenue_pie_data'),
    path('approve-request/<uuid:request_id>/', approve_request, name='approve_request'),
    path('reject-request/<uuid:request_id>/', reject_request, name='reject_request'),
    path('editHousehold/<uuid:household_id>/', editHousehold, name='edit_household'),
    path('deleteHousehold/<uuid:household_id>/', deleteHousehold, name='delete_household'),
    path('editResident/<uuid:member_id>/', editResident, name='edit_resident'),
    path('deleteResident/<uuid:member_id>/', deleteResident, name='delete_resident'),
    path('addhouseholds/', add_household, name='add_household'),
    path("users/", get_users),
    path('addresident/', add_resident, name='add_resident'),
    path('household/<uuid:household_id>/', get_household_info),
    path('deleteMember/<uuid:member_id>/', delete_member, name='delete_member'),
    path('recent-activities/', get_recent_activities, name='recent_activities'),
]

