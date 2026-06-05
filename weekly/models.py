from __future__ import annotations

from django.conf import settings
from django.db import models
from django.utils.text import slugify


class Employee(models.Model):
    payroll_number = models.PositiveIntegerField(unique=True, db_index=True)
    prox_id = models.PositiveIntegerField(null=True, blank=True)
    sage_pay_ref = models.PositiveIntegerField(null=True, blank=True)
    full_name = models.CharField(max_length=255)
    clock_name = models.CharField(max_length=255, blank=True)
    slug = models.SlugField(max_length=280, unique=True)
    company = models.CharField(max_length=255, blank=True)
    group = models.CharField(max_length=255, blank=True, db_index=True)
    status = models.CharField(max_length=64, blank=True, db_index=True)
    contract_hours = models.DecimalField(max_digits=8, decimal_places=2, default=0)
    basic_rate = models.DecimalField(max_digits=10, decimal_places=2, default=0)
    holidays_left = models.DecimalField(max_digits=8, decimal_places=2, null=True, blank=True)
    sage_pay_freq = models.CharField(max_length=64, blank=True)
    sage_paylink = models.CharField(max_length=128, blank=True)
    start_date = models.DateField(null=True, blank=True)
    finish_date = models.DateField(null=True, blank=True)
    shift_group = models.CharField(max_length=128, blank=True)
    ot_rule = models.CharField(max_length=128, blank=True)
    default_activity = models.CharField(max_length=128, blank=True)
    default_job = models.CharField(max_length=255, blank=True)
    notes = models.TextField(blank=True)
    is_active = models.BooleanField(default=True, db_index=True)
    imported_at = models.DateTimeField(auto_now=True)
    source_filename = models.CharField(max_length=255, blank=True)

    class Meta:
        ordering = ["full_name"]
        indexes = [
            models.Index(fields=["full_name"]),
        ]

    def __str__(self) -> str:
        return f"{self.full_name} ({self.payroll_number})"

    @staticmethod
    def make_slug(full_name: str, payroll_number: int) -> str:
        base = slugify(full_name) or "employee"
        return f"{base}-{payroll_number}"[:280]

    def save(self, *args, **kwargs) -> None:
        if not self.slug:
            self.slug = self.make_slug(self.full_name, self.payroll_number)
        super().save(*args, **kwargs)


class EmployeeDocument(models.Model):
    class DocType(models.TextChoices):
        HOLIDAY_FORM = "holiday_form", "Holiday form"
        CORRESPONDENCE = "correspondence", "Correspondence"
        OTHER = "other", "Other"

    employee = models.ForeignKey(
        Employee,
        on_delete=models.CASCADE,
        related_name="documents",
    )
    doc_type = models.CharField(max_length=32, choices=DocType.choices, default=DocType.OTHER)
    title = models.CharField(max_length=255)
    notes = models.TextField(blank=True)
    file = models.FileField(upload_to="employee_documents/%Y/%m/", blank=True)
    s3_key = models.CharField(max_length=512, blank=True)
    uploaded_at = models.DateTimeField(auto_now_add=True)
    uploaded_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="uploaded_employee_documents",
    )

    class Meta:
        ordering = ["-uploaded_at"]

    def __str__(self) -> str:
        return f"{self.title} ({self.employee_id})"
