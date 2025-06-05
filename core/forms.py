from django import forms
from .models import *
from datetime import datetime

class FeeForm(forms.ModelForm):
    class Meta:
        model = Fee
        fields = ['title', 'description', 'type', 'amount', 'due_date']

