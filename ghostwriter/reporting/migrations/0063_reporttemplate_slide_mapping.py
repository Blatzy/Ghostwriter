# Generated migration for adding slide_mapping to ReportTemplate

from django.db import migrations, models


class Migration(migrations.Migration):
    dependencies = [
        ("reporting", "0062_reporttemplate_contains_bloodhound_data"),
    ]

    operations = [
        migrations.AddField(
            model_name="reporttemplate",
            name="slide_mapping",
            field=models.JSONField(
                blank=True,
                null=True,
                help_text="Configuration for slide types, layouts, and ordering in PPTX templates",
                verbose_name="Slide Mapping Configuration",
            ),
        ),
    ]
