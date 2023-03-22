# -*- coding: utf-8 -*-
from django.core.management.base import BaseCommand, CommandError
from django.apps import apps

class Command(BaseCommand):
    help = "Reindex a given models."
    args = 'app.Model'

    def add_arguments(self, parser):
        parser.add_argument('models', nargs='+', type=str)

    def handle(self, *args, **options):
        models = options.get('models', None)
        if not models:
            raise CommandError('Please specify the model to reindex.')
        for model_name in models:
            parts = model_name.split('.', 1)
            # Model given is malformed
            if len(parts) != 2:
                raise CommandError('Indicate the model to reindex by following the syntax "app.Model".')
            # Get the model
            model = apps.get_model(*parts)
            # Callable model
            if model == None:
                raise CommandError('Unable to load the model "%s"' % model_name)
            saved_ct = 0
            self.stdout.write(self.style.SUCCESS('Starting reindex. This can take a while....'))
            # Load every objects to reindex them one by one
            for o in model.objects.all():
                # Save the object without changing anything will force a reindex
                o.save()
                # Count saved objects
                saved_ct += 1
            self.stdout.write(self.style.SUCCESS('Model "%s" reindexed through %s object(s).' % (model.__name__, saved_ct)))