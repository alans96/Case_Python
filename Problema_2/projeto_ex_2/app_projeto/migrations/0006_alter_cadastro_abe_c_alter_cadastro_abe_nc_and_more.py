# Generated by Django 5.0.1 on 2024-01-31 02:19

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('app_projeto', '0005_cadastro_abe_c_cadastro_abe_nc_cadastro_fec_c_and_more'),
    ]

    operations = [
        migrations.AlterField(
            model_name='cadastro',
            name='abe_c',
            field=models.CharField(blank=True, max_length=100, null=True),
        ),
        migrations.AlterField(
            model_name='cadastro',
            name='abe_nc',
            field=models.CharField(blank=True, max_length=100, null=True),
        ),
        migrations.AlterField(
            model_name='cadastro',
            name='fec_c',
            field=models.CharField(blank=True, max_length=100, null=True),
        ),
        migrations.AlterField(
            model_name='cadastro',
            name='fec_nc',
            field=models.CharField(blank=True, max_length=100, null=True),
        ),
        migrations.AlterField(
            model_name='cadastro',
            name='vor_valor',
            field=models.CharField(blank=True, max_length=100, null=True),
        ),
    ]
