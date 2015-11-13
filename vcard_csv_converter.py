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
from parser_exceptions import ParserException


class VcfToCsvConverter:
    def __init__(self, vcard_source, delimiter=",", quote=True):
        """delimeter:
            delimeter="," or delimeter=";"
        """

        if not vcard_source:
            raise ParserException('Empty vCard file')

        self.address_count = {'Home': 1, }
        self.telephone_count = {'Home Phone': 1,
                               'Work Phone': 1,
                               'Cell Phone': 1,
                               'Fax': 1}
        self.email_count = {
            'Personal' : 1,
            'Work' : 1
        }
        self.data = {}
        self.quote = quote
        self.delimiter = delimiter
        self.output = ''
        self.max_addresses = 1
        self.max_telephones = 1
        self.max_emails = 1
        self.vcard_source = vcard_source.split('\n')
        self.columns = (
            'Name', 'Organisation', 'Job Title',
            'Home Phone',
            'Work Phone',
            'Cell Phone',
            'Fax',
            'Personal Email',
            'Work Email',
            'Street Name', 'City', 'State/Province', 'Zip/Post Code', 'Country',
            'facebook', 'linkedin', 'twitter',
            'website',
        )

        self.data = self.__reset_row()
        for k in self.columns:
            self.__output(k)

        self.output += "\r\n"
        self.__parse_file()

    def __output_quote(self):
        if self.quote == True:
            self.output += '"'

    def __clean_data(self,text):
        if text[-1:] == ';' or text[-1:] == '\\':
            return self.__clean_data(text[:-1])
        return text

    def __output(self, text):
        self.__output_quote()
        text = text.replace('\\:',':').replace('\\;',';').replace('\\,',',').replace('\\=','=')
        if self.quote == True:
            text = text.replace('\\n','\n').replace('\\r','\r').replace('\\'+self.delimiter,self.delimiter).strip()
        else:
            text = text.strip()
        self.output += self.__clean_data(text)
        self.__output_quote()
        self.output += self.delimiter

    def __reset_row(self):
        self.address_count = {'Home': 1}
        self.telephone_count = {'Home Phone': 1,
                               'Work Phone': 1,
                               'Cell Phone': 1,
                               'Fax': 1, }
        self.email_count = {
            'Personal': 1,
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

    def __end_line(self):
        for k in self.columns:
            try:
                self.__output(self.data[ k ])
            except KeyError:
                self.output += self.delimiter

        self.output += "\r\n"
        self.data = self.__reset_row()

    def __parse_file(self):
        for line in self.vcard_source:
            self.__parse_line(line)

    def __parse_line(self, theLine):
        theLine = theLine.strip()
        if len(theLine) < 1:
            pass
        elif re.match('^BEGIN:VCARD', theLine, re.I):
            pass
        elif re.match('^END:VCARD', theLine, re.I):
            self.__end_line()
        else:
            self.__process_line(re.split("(?<!\\\\):", theLine))

    def __process_line(self, pieces):
        pre = re.split("(?<!\\\\);", pieces[0])
        if re.match('item.*', pre[0].split(".")[0], re.I) != None:
            try:
                pre[0] = pre[0].split(".")[1]
            except IndexError:
                print pre[0].split(".")

        if pre[0].upper() == 'N':
            self.__process_name(pieces[1])
        elif pre[0].upper() == 'FN':
            self.__process_single_value('Name', pieces[1])
        elif pre[0].upper() == 'TITLE':
            self.__process_single_value('Job Title', pieces[1])
        elif pre[0].upper() == 'ORG':
            self.__process_single_value('Organisation', pieces[1])
        elif pre[0].upper() == 'ADR':
            self.__process_address(pieces[1])
        elif pre[0].upper() == 'TEL':
            self.__process_telephone(pre, pieces[1:])
        elif pre[0].upper() == 'EMAIL':
            self.__process_email(pre, pieces[1:])
        elif pre[0].upper() == 'URL':
            self.__process_single_value('website', ":".join(pieces[1:]))

        self.data['facebook'] = ''
        self.data['linkedin'] = ''
        self.data['twitter'] = ''

    def __process_email(self, pre, p):
        hwm = "Personal"
        if re.search('work',(",").join(pre[1:]), re.I) != None:
            hwm = "Work"
        if self.email_count[hwm] <= self.max_emails:
            self.data["%s Email" % hwm] = p[0]

    def __process_telephone(self, pre, p):
        telephone_type = "Phone"
        hwm = "Home"
        if re.search('work', (",").join(pre[1:]), re.I) != None:
            hwm = "Work"
        elif re.search('cell', (",").join(pre[1:]), re.I) != None:
            hwm = "Cell"

        if re.search('fax', (",").join(pre[1:]), re.I) != None:
            telephone_type = "Fax"

        if telephone_type == 'Fax' and self.telephone_count[telephone_type] <= self.max_telephones:
            self.data[telephone_type] = p[0]
            self.telephone_count[telephone_type] += 1

        elif self.telephone_count[("%s Phone" % hwm)] <= self.max_telephones:
            self.data["%s %s" % (hwm, telephone_type)] = p[0]
            self.telephone_count["%s %s" % (hwm, telephone_type)] += 1

    def __process_address(self, p):
        try:
            (a, b, address, city, state, zip_code, country ) = re.split("(?<!\\\\);", p)
        except ValueError:
            (a, b, address, city, state, zip_code ) = re.split("(?<!\\\\);", p)
            country = ''

        address_type = "Home"

        if self.address_count[address_type] <= self.max_addresses:
            self.data["Street Name"] = address.strip()
            self.data["City"] = city.strip()
            self.data["State/Province"] = state.strip()
            self.data["Zip/Post Code"] = zip_code.strip()
            self.data["Country"] = country.strip()
            self.address_count[address_type] += 1

    def __process_single_value(self, value_type, p):
        self.data[value_type] = p

    def __process_name(self, p):
        try:
            (ln, fn, mi, pr, po) = re.split("(?<!\\\\);", p)
            self.data['Name'] = '%s %s %s %s %s' % (pr.strip(), fn.strip(), mi.strip(), ln.strip(), po.strip())
        except ValueError:
            if not self.data['Name'] is None:
                self.data['Name'] = p.strip()
