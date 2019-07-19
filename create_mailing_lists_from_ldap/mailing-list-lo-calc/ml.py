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
from sys import argv, stdout, stderr, exit
from os.path import dirname, join, abspath
from random import random
from os import kill
from signal import SIGKILL

if __name__=="__main__":
    if len(argv) < 2:
        print >>stderr, "\nFalta un argumento: por favor, especifique el archivo ODS\n"
        exit(1)

    try:
        import openoffice
        import openoffice.interact
        import openoffice.officehelper as OH
    except ImportError, e:
        print >>stderr, "Can't import openoffice-python module 'openoffice.interact'."
        if openoffice:
            print 'openoffice.__file__ :', openoffice.__file__
        exit(1)

    pipename = "uno" + str(random())[2:]
    connectString = OH._build_connect_string(pipename=pipename)
    pid = OH._start_OOo(connectString)
    try:
        context = OH.connect(connectString)
    except OH.BootstrapException:
        print >>stderr, "Can not connect to OOo. Please check whether it is running at all"
        exit(2)
    
    if not context:
        print >>stderr, "Can not connect to OOo (no context returned)."
        exit(2)

    smgr = context.ServiceManager
    desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",context)
    url = "file://" + abspath(argv[1])
    document = desktop.loadComponentFromURL(url, "_blank", 0, (openoffice.PropertyValue(),))
    try:
        sheets = document.getSheets()
    except Exception:
        raise TypeError("Model retrived was not a spreadsheet")
    sheet1 = getattr(sheets, sheets.ElementNames[0])
    print ("...\n\n")
    no_mail = []
    duplicated = []
#    ldap = ldap_open("ldap.estudiantes.uci.cu")
    ldap = ldap_open("ldap.uci.cu")
    i = 0
    student_name = unicode(sheet1.getCellByPosition(0, i).getString())
    while student_name != "":
        try:
            cn = student_name.encode("ascii", "replace").replace("?", "*")
            id = ldap.search("", SCOPE_SUBTREE, "cn=%s" % cn, ["mail"])
            st, data = ldap.result(id, 10)
            if len(data) < 1:
                no_mail.append(student_name)
            elif len(data) > 1:
                duplicated.append(student_name)
            else:
                line = "%s <%s>" % (student_name, data[0][1]["mail"][0])
                line = line.encode("utf-8")
                print line
        except:
            no_mail.append(student_name)
        i += 1
        student_name = unicode(sheet1.getCellByPosition(0, i).getString())
    kill(pid, SIGKILL)        
    if len(no_mail) > 0:
        print >>stderr, "\nADVERTENCIA: No se pudo encontrar el e-mail de los siguentes estudiantes:"
        for name in no_mail:
            print >>stderr, name
    if len(duplicated) > 0:
        print >>stderr, "\nADVERTENCIA: Los siguientes nombres están repetidos en el directorio (añada sus correos manualmente):"
        for name in duplicated:
            print >>stderr, name
