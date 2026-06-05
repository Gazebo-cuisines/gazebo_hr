from django.contrib import admin

from .models import Employee, EmployeeDocument


class EmployeeDocumentInline(admin.TabularInline):
    model = EmployeeDocument
    extra = 0
    fields = ("doc_type", "title", "file", "s3_key", "uploaded_at", "uploaded_by")
    readonly_fields = ("uploaded_at",)


@admin.register(Employee)
class EmployeeAdmin(admin.ModelAdmin):
    list_display = (
        "full_name",
        "payroll_number",
        "group",
        "status",
        "contract_hours",
        "is_active",
        "imported_at",
    )
    list_filter = ("is_active", "status", "group")
    search_fields = ("full_name", "clock_name", "payroll_number", "prox_id", "slug")
    readonly_fields = ("slug", "imported_at")
    inlines = [EmployeeDocumentInline]


@admin.register(EmployeeDocument)
class EmployeeDocumentAdmin(admin.ModelAdmin):
    list_display = ("title", "employee", "doc_type", "uploaded_at", "uploaded_by")
    list_filter = ("doc_type",)
    search_fields = ("title", "employee__full_name", "employee__payroll_number")
    raw_id_fields = ("employee",)
