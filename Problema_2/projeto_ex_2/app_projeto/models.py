from django.db import models

class Cadastro(models.Model):
    cas = models.CharField(max_length=255, null=True)
    contaminante = models.CharField(max_length=255, null=True)
    vor = models.CharField(max_length=100, blank=True, null=True)
    vor_valor = models.CharField(max_length=100, blank=True, null=True)
    abe_c = models.CharField(max_length=100, blank=True, null=True)
    abe_nc = models.CharField(max_length=100, blank=True, null=True)
    fec_c = models.CharField(max_length=100, blank=True, null=True)
    fec_nc = models.CharField(max_length=100, blank=True, null=True)