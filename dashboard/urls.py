# dashboard/urls.py
from django.urls import path
from . import views

app_name = 'dashboard'

urlpatterns = [
    # ============================================
    # AUTHENTICATION
    # ============================================
    path('', views.dashboard, name='dashboard'),
    path('login/', views.login_view, name='login'),
    path('logout/', views.logout_view, name='logout'),
    
    # ============================================
    # API ENDPOINTS - WEIGHT (REALTIME)
    # ============================================
    path('api/latest/', views.api_latest, name='api_latest'),
    path('api/sync/', views.api_sync, name='api_sync'),
    path('get-latest-weight/', views.get_latest_weight, name='get_latest_weight'),
    path('sync-weight-data/', views.sync_weight_data, name='sync_weight_data'),
    path('get-weight-history/', views.get_weight_history, name='get_weight_history'),
    path('get-thingspeak-data/', views.get_thingspeak_data, name='get_thingspeak_data'),
    path('test-api-connection/', views.test_api_connection, name='test_api_connection'),
    
    # ============================================
    # DATA MANAGEMENT
    # ============================================
    path('data-barang/', views.data_barang, name='data_barang'),
    path('data-kustomer/', views.data_kustomer, name='data_kustomer'),
    path('data-supplier/', views.data_supplier, name='data_supplier'),
    
    # ============================================
    # EXPORT/IMPORT
    # ============================================
    path('export-barang-excel/', views.export_barang_excel, name='export_barang_excel'),
    path('import-barang-excel/', views.import_barang_excel, name='import_barang_excel'),
    path('export-kustomer-excel/', views.export_kustomer_excel, name='export_kustomer_excel'),
    path('import-kustomer-excel/', views.import_kustomer_excel, name='import_kustomer_excel'),
    path('export-supplier-excel/', views.export_supplier_excel, name='export_supplier_excel'),
    path('import-supplier-excel/', views.import_supplier_excel, name='import_supplier_excel'),
    
    # ============================================
    # REPORT
    # ============================================
    path('report/', views.report, name='report'),
    path('report/detail/<int:pk>/', views.report_detail, name='report_detail'),
    path('report/delete/<int:pk>/', views.report_delete, name='report_delete'),
    path('report/update/<int:pk>/', views.report_update, name='report_update'),
    path('report/export-excel/', views.export_report_excel, name='export_report_excel'),
    path('report/export-selected/', views.export_selected_excel, name='export_selected_excel'),
    path('report/print/', views.print_report, name='print_report'),
    path('report/delete-selected/', views.delete_selected_records, name='delete_selected_records'),
    path('report/delete-all/', views.delete_all_records, name='delete_all_records'),
    path('report/delete-item/<int:item_id>/', views.delete_report_item, name='delete_report_item'),
    path('upload-report-photo/', views.upload_report_photo, name='upload_report_photo'),
    
    # ============================================
    # TRANSACTION
    # ============================================
    path('save-transaction/', views.save_transaction, name='save_transaction'),
    path('save-capture/', views.save_capture, name='save_capture'),
    
    # ============================================
    # SETTINGS
    # ============================================
    path('setting/', views.setting, name='setting'),
    
    # ============================================
    # PRINT TICKET
    # ============================================
    path('print-ticket/', views.print_ticket, name='print_ticket'),
    path('detect-printers/', views.detect_printers, name='detect_printers'),
    
    # ============================================
    # USER MANAGEMENT (FDA 21 CFR Part 11)
    # ============================================
    path('user-management/', views.user_management, name='user_management'),
    
    # ✅ USER CRUD
    path('users/add/', views.add_user, name='add_user'),
    path('users/delete/', views.delete_user, name='delete_user'),
    path('users/delete-all/', views.delete_all_users, name='delete_all_users'),
    path('users/export-excel/', views.export_users_excel, name='export_users_excel'),
    path('users/export-audit-csv/', views.export_audit_csv, name='export_audit_csv'),
    
    # ✅ GET & UPDATE USER
    path('dashboard/get-user/<str:username>/', views.get_user, name='get_user'),
    path('dashboard/update-user/', views.update_user, name='update_user'),
    
    # ✅ AUDIT LOG
    path('api/audit/delete/', views.delete_audit_log, name='delete_audit_log'),
    path('api/audit/delete-all/', views.delete_all_audit_logs, name='delete_all_audit_logs'),
    
    # ============================================
    # CRUD OPERATIONS - BARANG
    # ============================================
    path('get-barang/<str:id>/', views.get_barang, name='get_barang'),
    path('update-barang/', views.update_barang, name='update_barang'),
    path('delete-barang/', views.delete_barang, name='delete_barang'),
    
    # ============================================
    # CRUD OPERATIONS - KUSTOMER
    # ============================================
    path('get-kustomer/<str:id>/', views.get_kustomer, name='get_kustomer'),
    path('update-kustomer/', views.update_kustomer, name='update_kustomer'),
    path('delete-kustomer/', views.delete_kustomer, name='delete_kustomer'),
    
    # ============================================
    # CRUD OPERATIONS - SUPPLIER
    # ============================================
    path('get-supplier/<str:id>/', views.get_supplier, name='get_supplier'),
    path('update-supplier/', views.update_supplier, name='update_supplier'),
    path('delete-supplier/', views.delete_supplier, name='delete_supplier'),
]