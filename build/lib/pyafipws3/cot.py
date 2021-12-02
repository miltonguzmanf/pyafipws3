#!/usr/bin/python
# -*- coding: latin-1 -*-
# This program is free software; you can redistribute it and/or modify
# it under the terms of the GNU Lesser General Public License as published by the
# Free Software Foundation; either version 3, or (at your option) any later
# version.
#
# This program is distributed in the hope that it will be useful, but
# WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTIBILITY
# or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License
# for more details.

# Based on MultipartPostHandler.py (C) 02/2006 Will Holcomb <wholcomb@gmail.com>
# Ejemplos iniciales gracias a "Matias Gieco matigro@gmail.com"

"Módulo para obtener remito electrónico automático (COT)"

__author__ = "Mariano Reingart (reingart@gmail.com)"
__copyright__ = "Copyright (C) 2010 Mariano Reingart"
__license__ = "LGPL 3.0"
__version__ = "1.02h"

import os, sys, traceback
import base64
from pysimplesoap.simplexml import SimpleXMLElement
from io import StringIO
from .utils import WebClient, FileBufferString
from datetime import datetime

HOMO = False
CACERT = "conf/arba.crt"  # establecimiento de canal seguro (en producción)

##URL = "https://cot.ec.gba.gob.ar/TransporteBienes/SeguridadCliente/presentarRemitos.do"
# Nuevo servidor para el "Remito Electrónico Automático"
URL = "http://cot.test.arba.gov.ar/TransporteBienes/SeguridadCliente/presentarRemitos.do"  # testing


# URL = "https://cot.arba.gov.ar/TransporteBienes/SeguridadCliente/presentarRemitos.do"  # prod.


class COT:
    "Interfaz para el servicio de Remito Electronico ARBA"
    _public_methods_ = ['Conectar', 'PresentarRemito', 'LeerErrorValidacion',
                        'LeerValidacionRemito',
                        'AnalizarXml', 'ObtenerTagXml']
    _public_attrs_ = ['Usuario', 'Password', 'XmlResponse',
                      'Version', 'Excepcion', 'Traceback', 'InstallDir',
                      'CuitEmpresa', 'NumeroComprobante', 'CodigoIntegridad', 'NombreArchivo',
                      'TipoError', 'CodigoError', 'MensajeError',
                      'NumeroUnico', 'Procesado',
                      ]

    _reg_progid_ = "COT"
    _reg_clsid_ = "{7518B2CF-23E9-4821-BC55-D15966E15620}"

    Version = "%s %s" % (__version__, HOMO and 'Homologación' or '')

    def __init__(self):
        self.Usuario = self.Password = self.Token = self.Sign = None
        self.TipoError = self.CodigoError = self.MensajeError = ""
        self.LastID = self.LastCMP = self.CAE = self.CAEA = self.Vencimiento = ''
        self.InstallDir = INSTALL_DIR
        self.client = None
        self.xml = None
        self.RemitoBase64 = None
        self.RemitoXMLBase64 = None
        self.RemitoFileBufferString = None
        self.ErrorValidacionesRemitos = False
        self.limpiar()

    def limpiar(self):
        self.remitos = []
        self.errores = []
        self.XmlResponse = ""
        self.Excepcion = self.Traceback = ""
        self.TipoError = self.CodigoError = self.MensajeError = ""
        self.CuitEmpresa = self.NumeroComprobante = ""
        self.NombreArchivo = self.CodigoIntegridad = ""
        self.NumeroUnico = self.Procesado = ""
        self.XmlResponseBase64 = None

    def Conectar(self, url=None, proxy="", wrapper=None, cacert=None, trace=False, wsaa=None, user=None, password=None):
        if HOMO or not url:
            url = URL
        self.client = WebClient(location=url, trace=trace, cacert=cacert)
        self.wsaa = wsaa
        self.Usuario = user
        self.Password = password

    def setWsaa(self, wsaa):
        self.wsaa = wsaa

    def PresentarRemito(self, filename=None, testing=""):
        self.limpiar()
        try:
            if filename:
                if not os.path.exists(filename):
                    self.Excepcion = "Archivo no encontrado: %s" % filename
                    return False

                archivo = open(filename, "r")
            elif self.RemitoFileBufferString:
                archivo = self.RemitoFileBufferString
            else:
                self.Excepcion = "No existe archivo a enviar."
                return False
            if not testing:
                if self.wsaa:
                    content = self.client(token=self.wsaa.token, sign=self.wsaa.sign, file=archivo)
                    if self.client.response.status != 200:
                        self.Excepcion = "AGIP\nEstado: %s.\nRespuesta: %s." % (
                            self.client.response.reason, self.client.response.status
                        )
                        return False
                else:
                    content = self.client(user=self.Usuario, password=self.Password, file=archivo)
                    if self.client.response.status != 200:
                        self.Excepcion = "ARBA\nEstado: %s.\nRespuesta: %s." % (
                            self.client.response.reason, self.client.response.status
                        )
                        return False
            else:
                content = open(testing).read()
            self.XmlResponse = content
            self.xml = SimpleXMLElement(content)
            self.XmlResponseBase64 = base64.b64encode(content)
            if 'tipoError' in self.xml:
                self.TipoError = str(self.xml.tipoError)
                self.CodigoError = str(self.xml.codigoError)
                self.MensajeError = str(self.xml.mensajeError)
                self.ErrorValidacionesRemitos = True
            if 'cuitEmpresa' in self.xml:
                self.CuitEmpresa = str(self.xml.cuitEmpresa)
                self.NumeroComprobante = str(self.xml.numeroComprobante)
                self.NombreArchivo = str(self.xml.nombreArchivo)
                self.CodigoIntegridad = str(self.xml.codigoIntegridad)
                if 'validacionesRemitos' in self.xml:
                    for remito in self.xml.validacionesRemitos.remito:
                        d = {
                            'NumeroUnico': str(remito.numeroUnico),
                            'Procesado': str(remito.procesado),
                            'Errores': [],
                        }
                        if 'errores' in remito:
                            self.ErrorValidacionesRemitos = True
                            for error in remito.errores.error:
                                d['Errores'].append((
                                    str(error.codigo),
                                    str(error.descripcion)))
                        self.remitos.append(d)
                    # establecer valores del primer remito (sin eliminarlo)
                    self.LeerValidacionRemito(pop=False)
            return True
        except Exception as e:
            ex = traceback.format_exception(sys.exc_info()[0], sys.exc_info()[1], sys.exc_info()[2])
            self.Traceback = ''.join(ex)
            try:
                self.Excepcion = traceback.format_exception_only(sys.exc_info()[0], sys.exc_info()[1])[0]
            except:
                self.Excepcion = "<no disponible>"
            return False

    def LeerValidacionRemito(self, pop=True):
        "Leeo el próximo remito"
        # por compatibilidad hacia atras, la primera vez no remueve de la lista
        # (llamado de PresentarRemito con pop=False)
        if self.remitos:
            remito = self.remitos[0]
            if pop:
                del self.remitos[0]
            self.NumeroUnico = remito['NumeroUnico']
            self.Procesado = remito['Procesado']
            self.errores = remito['Errores']
            return True
        else:
            self.NumeroUnico = ""
            self.Procesado = ""
            self.errores = []
            return False

    def LeerErrorValidacion(self):
        if self.errores:
            error = self.errores.pop()
            self.TipoError = ""
            self.CodigoError = error[0]
            self.MensajeError = error[1]
            return True
        else:
            self.TipoError = ""
            self.CodigoError = ""
            self.MensajeError = ""
            return False

    def AnalizarXml(self, xml=""):
        "Analiza un mensaje XML (por defecto la respuesta)"
        try:
            if not xml:
                xml = self.XmlResponse
            self.xml = SimpleXMLElement(xml)
            return True
        except Exception as e:
            self.Excepcion = "%s" % (e)
            return False

    def ObtenerTagXml(self, *tags):
        "Busca en el Xml analizado y devuelve el tag solicitado"
        # convierto el xml a un objeto
        try:
            if self.xml:
                xml = self.xml
                # por cada tag, lo busco segun su nombre o posición
                for tag in tags:
                    xml = xml(tag)  # atajo a getitem y getattr
                # vuelvo a convertir a string el objeto xml encontrado
                return str(xml)
        except Exception as e:
            self.Excepcion = "%s" % (e)

    def crearRemito(self, kwargs):
        """
            # ver https://www.arba.gov.ar/archivos/Publicaciones/nuevodiseniodearchivotxt.pdf explicacion de los kwargs
            kwargs = {
                'HEADER': {
                    'TIPO_REGISTRO': '01',
                    'CUIT_EMPRESA': '30716396416',
                },
                'REMITO': {
                    'TIPO_REGISTRO': '02',
                    'REMITOS': [
                        {
                            'FECHA_EMISION': '20191014',  # formato YYYYMMDD,
                            'CODIGO_UNICO': '0910999990006816',
                            'FECHA_SALIDA_TRANSPORTE': '20191017',  # formato YYYYMMDD,
                            'SUJETO_GENERADOR': 'E',  # Valores posibles: E (emisor), D(destinatario),
                            'DESTINATARIO_CONSUMIDOR_FINAL': 0,  # I 0 = NO / 1 = SI,
                            'DESTINATARIO_TIPO_DOCUMENTO': '',
                            'DESTINATARIO_DOCUMENTO': '',
                            'DESTINATARIO_CUIT': '30682115722',  # Requerido si consumidor final=0 Si SUJETO_GENERADOR = 'D' debe ser igual a CUIT_EMPRESA
                            'DESTINATARIO_RAZON_SOCIAL': 'COMPUMUNDO S.A',  # Requerido si consumidor final = 0 Ver Nota campo DESTINATARIO_CONSUMIDOR_FINAL
                            'DESTINATARIO_TENEDOR': 0,
                            'DESTINO_DOMICILIO_CALLE': 'Villa Lugano',
                            'DESTINO_DOMICILIO_CODIGOPOSTAL': '1439',
                            'DESTINO_DOMICILIO_LOCALIDAD': 'Av. riestra',
                            'DESTINO_DOMICILIO_PROVINCIA': 'C',  # Ver Tabla de Provincias https://www.arba.gov.ar/archivos/Publicaciones/tablasdevalidacion.pdf
                            'ENTREGA_DOMICILIO_ORIGEN': 'NO',  # Valores: SI, NO
                            'ORIGEN_CUIT': '30716396416',  # Si SUJETO_GENERADOR = ?E? debe ser igual a CUIT_EMPRESA
                            'ORIGEN_RAZON_SOCIAL': 'COMPUMUNDO S.A.',
                            'ORIGEN_DOMICILIO_CALLE': 'San Mar n 5797',
                            'ORIGEN_DOMICILIO_CODIGOPOSTAL': '1766',
                            'ORIGEN_DOMICILIO_LOCALIDAD': 'TABLADA',
                            'ORIGEN_DOMICILIO_PROVINCIA': 'B',
                            'TRANSPORTISTA_CUIT': '20045162673',
                            'PATENTE_VEHICULO': '',  # 3 letras y 3 números ó 2 letras, 3 números y 2 letras. Requerido si TRANSPORTISTA_CUIT = CUIT_EMPRESA
                            'PRODUCTO_NO_TERM_DEV': 1,  # 0=NO / 1=SI
                            'IMPORTE': 1234,  # 12 enteros y 2 decimales. Las últimas 2 posiciones se toman como decimales siempre.Obligatorio siempre o sea > 0 (cero), salvo si campo PRODUCTO_NO_TERM_DEV = 1 o cuando ORIGEN_CUIT = DESTINATARIO_CUIT, que se permite 0 (cero) o blanco.
                            'PRODUCTOS': {
                                'TIPO_REGISTRO': '03',
                                'productos': [{
                                    'CODIGO_UNICO_PRODUCTO': '847150',  # longitud: 6 dígitos
                                    'ARBA_CODIGO_UNIDAD_MEDIDA': '3',   # ver tabla: https://www.arba.gov.ar/archivos/Publicaciones/tablasdevalidacion.pdf
                                    'CANTIDAD': '100', # 13 enteros y 2 decimales. > 0. Valores decimales son obligatorios. Ej: si la CANTIDAD es 200, enviar el valor CANTIDAD=20000.
                                    'PROPIO_CODIGO_PRODUCTO': '23891',
                                    'PROPIO_DESCRIPCION_PRODUCTO': 'COMP. SP-3960 VP',
                                    'PROPIO_DESCRIPCION_UNIDAD_MEDIDA': '1',
                                    'CANTIDAD_AJUSTADA': 100  # 13 enteros y 2 decimales. > 0. Valores decimales son obligatorios. Ej: si la CANTIDAD es 200, enviar el valor CANTIDAD=20000.
                                }]
                            }
                        }
                    ]
                },
                'FOOTER': {
                    'TIPO_REGISTRO': '04',
                    'CANTIDAD_TOTAL_REMITOS': 1
                },
                'NOMBREREMITOTXT': {
                    'CUIT_EMPRESA': '30716396416',
                    'NUM_PLANTA': '000',  # 3 digitos
                    'NUM_PUERTA': '000',  # 3 digitos
                    'FECHA_EMISION': '20191014',  # formato YYYYMMDD
                    'NUM_SECUENCIA': '000001',  # 6 digitos
                }
            }
        """
        kwargs_in_order = {
            'HEADER': ['CUIT_EMPRESA'],
            'REMITO': [
                'FECHA_EMISION', 'CODIGO_UNICO', 'FECHA_SALIDA_TRANSPORTE', 'HORA_SALIDA_TRANSPORTE',
                'SUJETO_GENERADOR',
                'DESTINATARIO_CONSUMIDOR_FINAL', 'DESTINATARIO_TIPO_DOCUMENTO', 'DESTINATARIO_DOCUMENTO',
                'DESTINATARIO_CUIT',
                'DESTINATARIO_RAZON_SOCIAL', 'DESTINATARIO_TENEDOR', 'DESTINO_DOMICILIO_CALLE',
                'DESTINO_DOMICILIO_NUMERO',
                'DESTINO_DOMICILIO_COMPLE', 'DESTINO_DOMICILIO_PISO', 'DESTINO_DOMICILIO_DTO',
                'DESTINO_DOMICILIO_BARRIO',
                'DESTINO_DOMICILIO_CODIGOPOSTAL', 'DESTINO_DOMICILIO_LOCALIDAD', 'DESTINO_DOMICILIO_PROVINCIA',
                'PROPIO_DESTINO_DOMICILIO_CODIGO', 'ENTREGA_DOMICILIO_ORIGEN', 'ORIGEN_CUIT', 'ORIGEN_RAZON_SOCIAL',
                'EMISOR_TENEDOR', 'ORIGEN_DOMICILIO_CALLE', 'ORIGEN_DOMICILIO_NUMERO ', 'ORIGEN_DOMICILIO_COMPLE',
                'ORIGEN_DOMICILIO_PISO', 'ORIGEN_DOMICILIO_DTO', 'ORIGEN_DOMICILIO_BARRIO',
                'ORIGEN_DOMICILIO_CODIGOPOSTAL',
                'ORIGEN_DOMICILIO_LOCALIDAD', 'ORIGEN_DOMICILIO_PROVINCIA', 'TRANSPORTISTA_CUIT', 'TIPO_RECORRIDO',
                'RECORRIDO_LOCALIDAD', 'RECORRIDO_CALLE', 'RECORRIDO_RUTA', 'PATENTE_VEHICULO', 'PATENTE_ACOPLADO',
                'PRODUCTO_NO_TERM_DEV', 'IMPORTE'],
            'PRODUCTOS': [
                'CODIGO_UNICO_PRODUCTO', 'ARBA_CODIGO_UNIDAD_MEDIDA', 'CANTIDAD', 'PROPIO_CODIGO_PRODUCTO',
                'PROPIO_DESCRIPCION_PRODUCTO', 'PROPIO_DESCRIPCION_UNIDAD_MEDIDA', 'CANTIDAD_AJUSTADA'
            ],
            'FOOTER': ['CANTIDAD_TOTAL_REMITOS'],
            'NOMBREREMITOTXT': ['CUIT_EMPRESA', 'NUM_PLANTA', 'NUM_PUERTA', 'FECHA_EMISION', 'NUM_SECUENCIA']
        }

        header_mandatory_fields = ['HEADER', 'REMITO', 'FOOTER', 'NOMBREREMITOTXT']
        for f in header_mandatory_fields:
            if f not in kwargs:
                self.Excepcion = "El diccionario 'kwargs' debe contener los elementos:\n%s " % '\n'.join(
                    header_mandatory_fields)
                return False

        remitos = kwargs['REMITO'].get('REMITOS', {})

        if not len(remitos):
            self.Excepcion = 'Debe existir al menos un remito.'

        # ver para el nombre de archivo: https://www.arba.gov.ar/archivos/Publicaciones/remito%20electr%C3%B3nico%20autom%C3%A1tico%20instructivo.pdf
        name = 'TB_%s_%s%s_%s_%s.txt' % tuple(
            kwargs['NOMBREREMITOTXT'].get(i, '') for i in kwargs_in_order['NOMBREREMITOTXT'])

        buf = StringIO()
        bufXML = StringIO()

        CUIT_EMPRESA = kwargs['HEADER'].get('CUIT_EMPRESA')
        buf.write('01|%s\n' % CUIT_EMPRESA)
        bufXML.write(
            '<?xml version\'1.0\' encoding=\'ISO-8859-1\'?>\n<HEADER>\n\t<CUIT_EMPRESA>%s</CUIT_EMPRESA>\n</HEADER>\n' %
            CUIT_EMPRESA
        )

        for remito in remitos:
            productos = remito['PRODUCTOS'].get('productos', [])

            # Begin lineas comentadas para manejarlas desde la respuesta del endpoint
            # if not len(productos):
            #     self.Excepcion = 'Debe existir al menos un producto en el remito a enviar.'
            #
            # cod_unico = remito['CODIGO_UNICO']
            # if not (remito.get('DESTINATARIO_CONSUMIDOR_FINAL') or len(remito.get('DESTINATARIO_CUIT', ''))):
            #     self.Excepcion = "Cuando DESTINATARIO_CONSUMIDOR_FINAL = 0, el campo DESTINATARIO_CUIT es obligatorio."
            #     return False
            #
            # if not (remito.get('DESTINATARIO_CONSUMIDOR_FINAL') or len(remito.get('DESTINATARIO_RAZON_SOCIAL', ''))):
            #     self.Excepcion = "Cuando DESTINATARIO_CONSUMIDOR_FINAL = 0, el campo 'DESTINATARIO_RAZON_SOCIAL es obligatorio."
            #     return False
            #
            # if remito.get('TRANSPORTISTA_CUIT') == CUIT_EMPRESA and not len(remito.get('PATENTE_VEHICULO')):
            #     self.Excepcion = "Si TRANSPORTISTA_CUIT = CUIT_EMPRESA el campo PATENTE_VEHICULO es obligatorio."
            #     return False
            #
            # if not remito.get('IMPORTE') and (not remito.get('PRODUCTO_NO_TERM_DEV') or CUIT_EMPRESA != remito.get('ORIGEN_CUIT')):
            #     self.Excepcion = "Si el campo PRODUCT_NO_TERM_DEV = 0 y CUITEMPRESA != ORIGEN_CUIT, el IMPORTE debe ser > 0."
            #     return False
            #
            # if remito.get('SUJETO_GENERADOR') == 'D' and CUIT_EMPRESA != remito.get('DESTINATARIO_CUIT'):
            #     self.Excepcion = "Si SUJETO_GENERADOR == 'D' el DESTINATARIO_CUIT debe ser igual a CUIT_EMPRESA."
            #     return False
            #
            # if remito.get('DESTINATARIO_CONSUMIDOR_FINAL') and remito.get('DESTINATARIO_TENEDOR'):
            #     self.Excepcion = "Si DESTINATARIO_CONSUMIDOR_FINAL == 1, DESTINATARIO_TENEDOR debe ser 0"
            #     return False
            #
            # if remito.get('SUJETO_GENERADOR') == 'E' and remito.get('ORIGEN_CUIT') != CUIT_EMPRESA:
            #     self.Excepcion = "Si SUJETO_GENERADOR == 'E' el ORIGEN_CUIT debe ser = CUIT_EMPRESA."
            #     return False
            # End lineas comentadas para manejarlas desde la respuesta del endpoint

            buf.write(('02' + '|%s' * len(kwargs_in_order['REMITO']) + '\n') %
                      tuple(str(remito.get(i, '')) for i in kwargs_in_order['REMITO']))

            def _get_tuple(array, obj):
                t = tuple()
                a = [(i, str(obj.get(i, '')), i) for i in array]
                for elem in a:
                    t += elem
                return t

            bufXML.write(('<REMITO>\n' + len(kwargs_in_order['REMITO']) * '\t<%s>%s</%s>\n') %
                         _get_tuple(kwargs_in_order['REMITO'], remito))

            for prod in productos:
                buf.write(('03' + len(kwargs_in_order['PRODUCTOS']) * '|%s' + '\n') %
                          tuple(str(prod.get(i, '')) for i in kwargs_in_order['PRODUCTOS']))

                bufXML.write(
                    ('\t<PRODUCTOS>\n' + len(kwargs_in_order['PRODUCTOS']) * '\t\t<%s>%s</%s>\n' + '\t</PRODUCTOS>\n') %
                    _get_tuple(kwargs_in_order['PRODUCTOS'], prod))

            bufXML.write('</REMITO>\n')

        kwargs['FOOTER']['CANTIDAD_TOTAL_REMITOS'] = len(remitos)

        buf.write('04|%s' % str(len(remitos)))

        bufXML.write('<FOOTER>%s</FOOTER>\n' % str(len(remitos)))

        content = buf.getvalue()
        buf.close()

        contentXML = bufXML.getvalue()
        bufXML.close()

        self.RemitoBase64 = base64.b64encode(bytes(content, 'utf-8'))
        self.RemitoXMLBase64 = base64.b64encode(bytes(contentXML, 'utf-8'))
        self.RemitoFileBufferString = FileBufferString(name, content)

        return True


# busco el directorio de instalación (global para que no cambie si usan otra dll)
if not hasattr(sys, "frozen"):
    basepath = __file__
elif sys.frozen == 'dll':
    import win32api

    basepath = win32api.GetModuleFileName(sys.frozendllhandle)
else:
    basepath = sys.executable
INSTALL_DIR = os.path.dirname(os.path.abspath(basepath))

if __name__ == "__main__":

    if "--register" in sys.argv or "--unregister" in sys.argv:
        import win32com.server.register

        win32com.server.register.UseCommandLine(COT)
        sys.exit(0)
    elif len(sys.argv) < 4:
        print("Se debe especificar el nombre de archivo, usuario y clave como argumentos!")
        sys.exit(1)

    cot = COT()
    filename = sys.argv[1]  # TB_20111111112_000000_20080124_000001.txt
    cot.Usuario = sys.argv[2]  # 20267565393
    cot.Password = sys.argv[3]  # 23456

    if '--testing' in sys.argv:
        test_response = "cot_response_multiple_errores.xml"
        # test_response = "cot_response_2_errores.xml"
        # test_response = "cot_response_3_sinerrores.xml"
    else:
        test_response = ""

    if not HOMO:
        for i, arg in enumerate(sys.argv):
            if arg.startswith("--prod"):
                URL = URL.replace("http://cot.test.arba.gov.ar",
                                  "https://cot.arba.gov.ar")
                print("Usando URL:", URL)
                break
            if arg.startswith("https"):
                URL = arg
                print("Usando URL:", URL)
                break

    cot.Conectar(URL, trace='--trace' in sys.argv, cacert=CACERT)
    cot.PresentarRemito(filename, testing=test_response)

    if cot.Excepcion:
        print("Excepcion:", cot.Excepcion)
        print("Traceback:", cot.Traceback)

    # datos generales:
    print("CUIT Empresa:", cot.CuitEmpresa)
    print("Numero Comprobante:", cot.NumeroComprobante)
    print("Nombre Archivo:", cot.NombreArchivo)
    print("Codigo Integridad:", cot.CodigoIntegridad)

    print("Error General:", cot.TipoError, "|", cot.CodigoError, "|", cot.MensajeError)

    # recorro los remitos devueltos e imprimo sus datos por cada uno:
    while cot.LeerValidacionRemito():
        print("Numero Unico:", cot.NumeroUnico)
        print("Procesado:", cot.Procesado)
        while cot.LeerErrorValidacion():
            print("Error Validacion:", "|", cot.CodigoError, "|", cot.MensajeError)

    # Ejemplos de uso ObtenerTagXml
    if False:
        print("cuit", cot.ObtenerTagXml('cuitEmpresa'))
        print("p0", cot.ObtenerTagXml('validacionesRemitos', 'remito', 0, 'procesado'))
        print("p1", cot.ObtenerTagXml('validacionesRemitos', 'remito', 1, 'procesado'))
