#!/usr/bin/env python
# -*- coding: utf-8 -*-
#       This program is free software; you can redistribute it and/or modify
#       it under the terms of the GNU General Public License as published by
#       the Free Software Foundation; either version 2 of the License, or
#       (at your option) any later version.
#       
#       This program is distributed in the hope that it will be useful,
#       but WITHOUT ANY WARRANTY; without even the implied warranty of
#       MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#       GNU General Public License for more details.
#       
#       You should have received a copy of the GNU General Public License
#       along with this program; if not, write to the Free Software
#       Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston,
#       MA 02110-1301, USA.
#
#  __autor__ = Dariem Pérez Herrera <dariemp en uci.cu>

from ldap import open as ldap_open, SCOPE_SUBTREE
from suds.client import Client
from suds.xsd.doctor import Import, ImportDoctor
from sys import argv, stderr
from os.path import dirname, join, abspath

if __name__=="__main__":
    if len(argv) < 2:
        print >>stderr, "\nFalta un argumento: por favor, especifique el grupo\n"
        exit(1)
    nombre_grupo = argv[1]
    xsd_loc = abspath(join(dirname(argv[0]), "schemas", "soap-encoding.xsd"))
    xsd_uri = "file://" + xsd_loc
    imp = Import("http://schemas.xmlsoap.org/soap/encoding/", location=xsd_uri)
    doc = ImportDoctor(imp)
    akademosWSDL = "http://akademos2.uci.cu/servicios/v3/AkademosWS.wsdl"
    client = Client(akademosWSDL, doctor = doc)
    grupos = client.service.ObtenerGrupos()
    i = 0
    grupo = None
    while i < len(grupos) and not grupo:
        if grupos[i].NombreGrupo == nombre_grupo:
            grupo = grupos[i]
        i += 1
    if not grupo:
        print >>stderr, "\nGrupo no encontrado: rectifique formato de entrada\n"
        exit(1)
    estudiantes = client.service.ObtenerEstudiantesDadoFiltro(grupo)
    no_mail = []
    duplicated = []
    ldap = ldap_open("ldap.uci.cu")
    for estudiante in estudiantes:
        unicode_cn = u"%s %s %s %s" % \
            (estudiante.PrimerNombre, estudiante.SegundoNombre,
             estudiante.PrimerApellido, estudiante.SegundoApellido)
        unicode_cn = unicode_cn.replace("   ", " ").replace(" - ", " ")
        cn = unicode_cn.encode("ascii", "replace").replace("?", "*")
        try:
            id = ldap.search("", SCOPE_SUBTREE, "cn=%s" % cn, ["mail"])
            st, data = ldap.result(id, 10)
            if len(data) < 1:
                no_mail.append(unicode_cn)
            elif len(data) > 1:
                duplicated.append(unicode_cn)
            else:
                line = "%s <%s>" % (unicode_cn, data[0][1]["mail"][0])
                line = line.encode("utf-8")
                print line
        except:
            no_mail.append(unicode_cn)
    if len(no_mail) > 0:
        print >>stderr, "\nADVERTENCIA: No se pudo encontrar el e-mail de los siguentes estudiantes:"
        for name in no_mail:
            print >>stderr, name
    if len(duplicated) > 0:
        print >>stderr, "\nADVERTENCIA: Los siguientes nombres están repetidos en el directorio (añada sus correos manualmente):"
        for name in duplicated:
            print >>stderr, name
