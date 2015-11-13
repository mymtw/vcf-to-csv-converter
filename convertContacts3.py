#!/usr/bin/python
# -*- coding: utf-8 -*-
"""
    VcfToCsvConverter v0.3 - Converts VCF/VCARD files into CSV
    Copyright (C) 2009 Petar Strinic (http://petarstrinic.com)
    Contributor -- Dave Dartt
    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.
    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
    GNU General Public License for more details.
    You should have received a copy of the GNU General Public License
    along with this program. If not, see <http://www.gnu.org/licenses/>.
"""
import re
from contact_import.parser_exceptions import ParserException


class VcfToCsvConverter:
    def __init__(self, vcard_source, delimiter=",", quote=True):
        """delimeter:
            delimeter="," or delimeter=";"
        """

        if not vcard_source:
            raise ParserException('Empty vCard file')

        self.addressCount = {'Home': 1,}
        self.telephoneCount = {'Home Phone': 1,
                               'Work Phone': 1,
                               'Cell Phone': 1,
                               'Fax': 1}
        self.emailCount = {
            'Home' : 1,
            'Work' : 1
        }
        self.data = {}
        self.quote = quote
        self.delimiter = delimiter
        self.output = ''
        self.maxAddresses = 1
        self.maxTelephones = 1
        self.maxEmails = 1
        self.vcard_source = vcard_source.split('\n')
        self.columns = (
            'Name', 'Organisation', 'Job Title',
            'Street Name', 'City', 'State/Province', 'Zip/Post Code', 'Country',
            'Home Phone',
            'Work Phone',
            'Cell Phone',
            'Fax',
            'Home Email',
            'Work Email',
            'website'
        )

        self.data = self.__resetRow()
        for k in self.columns:
            self.__output(k)

        self.output += "\r\n"
        self.__parseFile()

    def __outputQuote(self):
        if self.quote == True:
            self.output += '"'

    def __CleanData(self,text):
        if text[-1:] == ';' or text[-1:] == '\\':
            return self.__CleanData(text[:-1])
        return text

    def __output(self, text):
        self.__outputQuote();
        text = text.replace('\\:',':').replace('\\;',';').replace('\\,',',').replace('\\=','=')
        if self.quote == True:
            text = text.replace('\\n','\n').replace('\\r','\r').replace('\\'+self.delimiter,self.delimiter).strip()
        else:
            text = text.strip()
        self.output += self.__CleanData(text)
        self.__outputQuote()
        self.output += self.delimiter

    def __resetRow(self):
        self.addressCount = {'Home': 1}
        self.telephoneCount = {'Home Phone': 1,
                               'Work Phone': 1,
                               'Cell Phone': 1,
                               'Home Fax': 1,
                               'Work Fax': 1, }
        self.emailCount = {
            'Home': 1,
            'Work': 1
        }
        array = {}
        for k in self.columns:
            array[k] = ''
        return array

    def __setitem__(self, k, v):
        self.data[k] = v

    def __getitem__(self, k):
        return self.data[k]

    def __endLine(self):
        for k in self.columns:
            try:
                self.__output(self.data[ k ])
            except KeyError:
                self.output += self.delimiter

        self.output += "\r\n"
        self.data = self.__resetRow()

    def __parseFile(self):
        for line in self.vcard_source:
            self.__parseLine(line)

    def __parseLine(self, theLine):
        theLine = theLine.strip()
        if len(theLine) < 1:
            pass
        elif re.match('^BEGIN:VCARD', theLine, re.I):
            pass
        elif re.match('^END:VCARD', theLine, re.I):
            self.__endLine()
        else:
            self.__processLine(re.split("(?<!\\\\):", theLine))

    def __processLine(self, pieces):
        pre = re.split("(?<!\\\\);", pieces[0])
        if re.match('item.*', pre[0].split(".")[0], re.I) != None:
            try:
                pre[0] = pre[0].split(".")[1]
            except IndexError:
                print pre[0].split(".")

        if pre[0].upper() == 'N':
            self.__processName(pieces[1])
        elif pre[0].upper() == 'FN':
            self.__processSingleValue('Name', pieces[1])
        elif pre[0].upper() == 'TITLE':
            self.__processSingleValue('Job Title', pieces[1])
        elif pre[0].upper() == 'ORG':
            self.__processSingleValue('Organisation', pieces[1])
        elif pre[0].upper() == 'ADR':
            self.__processAddress(pieces[1])
        elif pre[0].upper() == 'TEL':
            self.__processTelephone(pre, pieces[1:])
        elif pre[0].upper() == 'EMAIL':
            self.__processEmail(pre, pieces[1:])
        elif pre[0].upper() == 'URL':
            self.__processSingleValue('website', ":".join(pieces[1:]))

    def __processEmail(self, pre, p):
        hwm = "Home"
        if re.search('work',(",").join(pre[1:]), re.I) != None:
            hwm = "Work"
        if self.emailCount[hwm] <= self.maxEmails:
            self.data["%s Email" % hwm] = p[0].capitalize()

    def __processTelephone(self, pre, p):
        telephoneType = "Phone"
        hwm = "Home"
        if re.search('work', (",").join(pre[1:]), re.I) != None:
            hwm = "Work"
        elif re.search('cell', (",").join(pre[1:]), re.I) != None:
            hwm = "Cell"

        if re.search('fax', (",").join(pre[1:]), re.I) != None:
            telephoneType = "Fax"

        if self.telephoneCount[("%s Phone" % hwm)] <= self.maxTelephones:
            self.data["%s %s" % (hwm, telephoneType)] = p[0].capitalize()
            self.telephoneCount["%s %s" % (hwm, telephoneType)] += 1

    def __processAddress(self, p):
        try:
            (a, b, address, city, state, zip_code, country ) = re.split("(?<!\\\\);", p)
        except ValueError:
            (a, b, address, city, state, zip_code ) = re.split("(?<!\\\\);", p)
            country = ''

        addressType = "Home"

        if self.addressCount[addressType] <= self.maxAddresses:
            self.data["Street Name"] = address.strip()
            self.data["City"] = city.strip()
            self.data["State/Province"] = state.strip()
            self.data["Zip/Post Code"] = zip_code.strip()
            self.data["Country"] = country.strip()
            self.addressCount[addressType] += 1

    def __processSingleValue(self, valueType, p):
        self.data[valueType] = p

    def __processName(self, p):
        try:
            (ln, fn, mi, pr, po) = re.split("(?<!\\\\);",p)
            self.data['Name'] = '%s %s %s %s %s' % (pr.strip(), fn.strip(), mi.strip(), ln.strip(), po.strip())
        except ValueError:
            if not self.data['Name'] is None:
                self.data['Name'] = p.strip()
