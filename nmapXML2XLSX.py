#!/usr/bin/python3

import io
import argparse
from lxml.etree import XMLParser, XML
from openpyxl import Workbook

try:
    from openpyxl.writer.excel import save_virtual_workbook
except ImportError:
    from openpyxl.writer.excel import ExcelWriter
    from zipfile import ZipFile, ZIP_DEFLATED

    def save_virtual_workbook(workbook):
        """Return an in-memory workbook"""
        temp_buffer = io.BytesIO()
        archive = ZipFile(temp_buffer, 'w', ZIP_DEFLATED, allowZip64=True)

        writer = ExcelWriter(workbook, archive)
        writer.save()

        virtual_workbook = temp_buffer.getvalue()
        temp_buffer.close()
        return virtual_workbook


class XMLParserTarget:
    def __init__(self):
        self.wb = Workbook(write_only=True)
        self._current_ws = self.wb.create_sheet(title='nmap output')
        self._metadata = ''
        self._free_host_valuables()

    def _free_host_valuables(self):
        self._current_host = None
        self._current_host_starttime = None
        self._current_host_endtime = None
        self._current_host_state = None
        self._current_host_hostnames = []
        self._port = None
        self._ports = []

    def start(self, tag, attrib):
        if tag == 'nmaprun':
            self._metadata = attrib['args']
        elif tag == 'times':
            # Reaching this tag means that we just finished parsing the whole host,
            # and it's time to write it down to the worksheet
            if not self._current_host:
                raise Exception('Somehow parsing went wrong with "times" tag')

            if self._current_host_state == 'up':
                for port in self._ports:
                    self._current_ws.append([self._current_host] + port)

            self._free_host_valuables()

        elif tag == 'host':
            # I don't use it, but maybe it will be in use later
            self._current_host_starttime = attrib['starttime']
            self._current_host_endtime = attrib['endtime']
        elif tag == 'status':
            self._current_host_state = attrib['state']
        elif tag == 'address':
            self._current_host = attrib['addr']
        elif tag == 'hostname':
            self._current_host_hostnames.append(attrib['name'])
        elif tag == 'port':
            """
            <port protocol="tcp" portid="80">
                <state state="open" reason="syn-ack" reason_ttl="44"/>
                <service name="http" method="table" conf="3"/>
            </port>
            """
            self._port = [attrib['protocol'] + '/' + attrib['portid']]
        elif tag == 'state':
            if not self._port:
                raise Exception('Somehow parsing went wrong with "state" tag')
            self._port.append(attrib['state'])
        elif tag == 'service':
            if not self._port:
                raise Exception('Somehow parsing went wrong with "service" tag')
            self._port.append(attrib['name'])
            self._ports.append(self._port)

            self._port = None

    def close(self):
        return save_virtual_workbook(self.wb)


def convert(xml_data):
    parser = XMLParser(target=XMLParserTarget(), encoding='UTF-8', remove_blank_text=True, huge_tree=True)
    return XML(xml_data, parser)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Convert nmap XML output to XLSX format')
    parser.add_argument('XML_FILE', help='nmap XML output')
    parser.add_argument('XLSX_FILE', help='name of the result XLSX file')
    args = parser.parse_args()

    with open(args.XML_FILE) as xml_file_obj:
        xml_data_lines = [x.replace('\n', '') for x in xml_file_obj.readlines()]

    if '<?xml version="1.0"' in xml_data_lines[0]:
        xml_data_lines.pop(0) # <?xml version="1.0" encoding="UTF-8"?>
    if '<!DOCTYPE' in xml_data_lines[0]:
        xml_data_lines.pop(0) # <!DOCTYPE nmaprun>
    if 'xml-stylesheet' in xml_data_lines[0]:
        xml_data_lines.pop(0)  # <?xml-stylesheet href="file:///usr/bin/../share/nmap/nmap.xsl" type="text/xsl"?>
    if '<!-- Nmap' in xml_data_lines[0]:
        xml_data_lines.pop(0) # <!-- Nmap 7.94SVN scan initiated Mon Sep 16 20:10:32 2024 as: nmap -oX default_1.nmap -iL s1.txt -Pn -->

    xml_data = '\n'.join(xml_data_lines)

    xlsx_data = convert(xml_data)

    with open(args.XLSX_FILE, "wb") as xlsx_file_obj:
        xlsx_file_obj.write(xlsx_data)
