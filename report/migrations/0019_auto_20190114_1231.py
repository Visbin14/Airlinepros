# Generated by Django 2.1.2 on 2019-01-14 12:31

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('report', '0018_auto_20190114_1227'),
    ]

    operations = [
        migrations.AlterField(
            model_name='transaction',
            name='balance',
            field=models.FloatField(default=0.0, null=True),
        ),
        migrations.AlterField(
            model_name='transaction',
            name='ca',
            field=models.FloatField(default=0.0, null=True),
        ),
        migrations.AlterField(
            model_name='transaction',
            name='cc',
            field=models.FloatField(default=0.0, null=True),
        ),
        migrations.AlterField(
            model_name='transaction',
            name='cobl_amount',
            field=models.FloatField(default=0.0, null=True),
        ),
        migrations.AlterField(
            model_name='transaction',
            name='ep',
            field=models.FloatField(default=0.0, null=True),
        ),
        migrations.AlterField(
            model_name='transaction',
            name='fare_amount',
            field=models.FloatField(default=0.0, null=True),
        ),
        migrations.AlterField(
            model_name='transaction',
            name='pen',
            field=models.FloatField(default=0.0, null=True),
        ),
        migrations.AlterField(
            model_name='transaction',
            name='std_comm_amount',
            field=models.FloatField(default=0.0, null=True),
        ),
        migrations.AlterField(
            model_name='transaction',
            name='std_comm_rate',
            field=models.FloatField(default=0.0, null=True),
        ),
        migrations.AlterField(
            model_name='transaction',
            name='sup_comm_amount',
            field=models.FloatField(default=0.0, null=True),
        ),
        migrations.AlterField(
            model_name='transaction',
            name='sup_comm_rate',
            field=models.FloatField(default=0.0, null=True),
        ),
        migrations.AlterField(
            model_name='transaction',
            name='tax_on_comm',
            field=models.FloatField(default=0.0, null=True),
        ),
    ]