#!/usr/bin/python3
import sys

from R2Log import logger
from rich.console import Console
from rich.table import Table
from rich.table import Column
from rich.progress import Progress, BarColumn, TextColumn
from lxml.builder import unicode
import json
from docxtpl import DocxTemplate

import os
import re
import glob
import html2text
import html

import argparse
import xml.etree.ElementTree as etree

from deep_translator import GoogleTranslator


VERSION = "1.0.0"

banner = """Generate CIS report from XML output - V%s \n\n/!\ \tPlease fill ./cis_xml/ with CIS-CAT XML output.\t /!\ \n""" % VERSION

group_title_depth = 1

def translate(sentence):
    sentence = html.unescape(sentence)
    sentence = sentence.replace("<x:", " ")
    sentence = sentence.replace("<", " ")
    sentence = sentence.replace(">", " ")
    sentence = sentence.replace("\\n", " ")
    translate = None
    try:
        translate = GoogleTranslator(source='auto', target='fr').translate(sentence)
    except:
        pass
    if translate is not None :
        logger.advanced(translate)
        return translate
    else:
        logger.error("Error translate.")
        logger.advanced(sentence)
        return sentence

class Entry(object):
    """Entry object, containing all the information for a control
        branch the name of the top-level branch for the control
        number the node number
        control the name of the control
        result the pass/fail/error result of the control
        description the description of the control
        remediation the remediation for the control
    """
    def __init__(self,
                 branch, number, control, result, description, remediation):
        logger.advanced(number)
        self.branch = translate(branch)
        self.number = number
        self.control = translate(control)
        if result == "notchecked":
            self.result = "manual"
        elif result == "notselected":
            self.result = "N/A"
        else:
            self.result = result
        self.description = translate(description)
        self.remediation = translate(remediation)

    def get_branch(self):
        return self.branch
    def get_number(self):
        return self.number
    def get_control(self):
        return self.control
    def get_result(self):
        return self.result
    def get_description(self):
        return self.description
    def get_remediation(self):
        return self.remediation

    def get_csv_string(self):
        # not sending description as it isnt perfect
        if self.result == 'pass':
            remediation_text = ''
        else:
            remediation_text = self.remediation
        return('"%s","%s","%s","%s"\n' % (
            self.branch,
            self.control,
            self.result,
            remediation_text))

def parseArgs():
    print(banner)
    parser = argparse.ArgumentParser(description="A python script to generate CIS report.")

    group_config = parser.add_argument_group("Configuration")
    group_config.add_argument("-v", "--verbose", action="count", default=0, help="Verbosity level (-v for verbose, -vv for advanced, -vvv for debug)")

    group_export = parser.add_argument_group("Export")
    group_export.add_argument("--export-docx", dest="export_docx", type=str, default=None, required=False,
                              help="Output DOCX file to store the results in.")

    args = parser.parse_args()

    if args.export_docx is None:
        logger.error("No DOCX specified.")
        sys.exit(1)

    return args

def description_node_to_text(node):
    """
        Takes an XML node containing nested strings and returns a single string.
    """

    h = html2text.HTML2Text()
    h.ignore_links = True
    html = etree.tostring(node)

    text = str(h.handle(str(html))
                .replace('b\'\\n ', '')
                .replace('\\n \\n \'', '')
                .replace('\\n \\n ', '')
                .replace('\\n ', '')
                .replace('\\n \\n', '')
                .replace('\\n', '')
                .replace('"', '\'')
                .replace('\\\'', '\'')
                .strip())

    # for some reason some text end with an '
    if text[-1] == '\'':
        text = text[:-2]

    return text


def remediation_node_to_text(node):
    """
        Takes an XML node containing nested strings and returns a single string.
    """

    h = html2text.HTML2Text()
    h.ignore_links = True
    html = etree.tostring(node)

    text = str(h.handle(str(html))
                .replace('b\'\\n ', '')
                .replace('\\n \\n \'', '')
                .replace('\\n \\n ', '')
                .replace('\\n ', '')
                .replace('\\n \\n', '')
                .replace('\\n', '')
                .replace('"', '\'')
                .replace('\\\'', '\'')
                .replace('\\\\\\', '\\')
                .strip())

    text = re.sub(r'\s+', ' ', text)

    text = text.split("Impact")[0]

    if text[0] == ' ':
        text = text[1:]
    if text[-1] == ' ':
        text = text[:-1]
    if text[-1] != '.':
        text += '.'

    text = text.replace('Computer Configuration', '\r\nComputer Configuration')

    return text


def recursive_iter_over_group(node, level):
    """
        As each group (branch) can contain either rules or sub-groups
        (sub-branches), we use a recursive function to iterate over each
        group/sub-group.
    """
    global group_title
    global group_title_depth


    for child in node:
        if 'title' in child.tag:
            if level == group_title_depth:
                group_title = child.text
        elif 'description' in child.tag:
            #    group_description = '-'
            #    group_description = recursive_get_string(child)
            pass
        elif 'Rule' in child.tag:
            rule_id = child.get('id')
            for i in child:
                if 'title' in i.tag:
                    #   rule_title = i.text.replace('\n            ', ' ')
                    rule_title = re.sub(r'\s+', ' ', i.text)
                elif 'description' in i.tag:
                    if len(i):
                        rule_description = description_node_to_text(i)
                    else:
                        rule_description = i.text
                elif 'fixtext' in i.tag:
                    rule_remediation = '-'
                    rule_remediation = remediation_node_to_text(i)

                rule_number = rule_id.split('_')[3]
                rule_result = result_dict[rule_id]


                if rule_number == "5.3.4":
                    pass
            # pass : OK
            # fail : KO
            # notchecked : manual check
            # notselected : N/A


            # TODO make sure all values are set
            if rule_result != 'notselected' and \
                    rule_result != 'unknown':
                new_entry = Entry(group_title,
                                  rule_number,
                                  rule_title,
                                  rule_result,
                                  rule_description,
                                  rule_remediation)
                entry_list.append(new_entry)
            else:
                pass

        elif 'Group' in child.tag:
            recursive_iter_over_group(child, level+1)
        else:  # unhandled case
            print('[-] Unhandled tag %s.' % child.tag)

def parse_cis_html(pathCisXmlFile, name):
    #from https://github.com/x4v13r64/CISCAT_xml2csv/blob/master/ciscat_xml2csv.py
    nameCisXmlFile = ''
    if os.name == "nt":
        nameCisXmlFile = pathCisXmlFile.split("\\")[-1]
        logger.info("Parsing " + nameCisXmlFile + " file...")
        tree = etree.parse(pathCisXmlFile)
        root = tree.getroot()

        """
        First step is to build a dict with Rule ids and pass/fail results
        XML structure:
        <Benchmark xmlns="http://checklists.nist.gov/xccdf/1.2"
           <TestResult end-time="2015-12-18T10:10:51.776+01:00">
                 <rule-result idref="xccdf_org.cisecurity.benchmarks_rule_1.1.1_L1_ ...
                 Set_Enforce_password_history_to_24_or_more_passwords"
        """
        global result_dict
        result_dict = dict()  # contains id-pass/fail/error/notselected results




        for reports in root:  # iterate over root
            for report in reports:
                for content in report:
                    for oval_results in content:
                            for i in oval_results:
                                if 'rule-result' in i.tag:  # each rule-result contains one result
                                    idref = i.get('idref')  # Rule id
                                    for j in i:
                                        if 'result' in j.tag:  # result
                                            result_dict[idref] = j.text
        """
                        for results in oval_results:
                            if 'TestResult' in results.tag:
                                logger.success("OK")
                            for system in results:
                                if 'TestResult' in system.tag:
                                    logger.success("OK")
                                for oval_system_characteristics in system:
                                    if 'TestResult' in oval_system_characteristics.tag:
                                        logger.success("OK")
                                    for system_data in oval_system_characteristics:
                                        if 'TestResult' in system_data.tag:
                                            logger.success("OK")
        """
        """
        for child in root:
            if 'TestResult' in child.tag:  # TestResult contains all the results
                for i in child:
                    if 'rule-result' in i.tag:  # each rule-result contains one result
                        idref = i.get('idref')  # Rule id
                        for j in i:
                            if 'result' in j.tag:  # result
                                result_dict[idref] = j.text
        """

        """
        Second step is to parse all Groups and their content, and cross-reference with
        the result dict for pass/fail/error/notselected result.
        XML structure:
        <Benchmark xmlns="http://checklists.nist.gov/xccdf/1.2"
            <Group id="xccdf_org.cisecurity.benchmarks_group_1_Account_Policies">
                <Group id="xccdf_org.cisecurity.benchmarks_group_1.1_Password_Policy">
                    <Rule id="xccdf_org.cisecurity.benchmarks_rule_1.1.1_L1_Set_ ...
                    Enforce_password_history_to_24_or_more_passwords">
        """
        global entry_list
        entry_list = []


        for child in root:  # iterate over root
            if 'Group' in child.tag:  # each Group is a branch
                recursive_iter_over_group(child, 0)  # recursive iterate over each group

        text_column = TextColumn("Parsing XML (2/3 minutes) ...", table_column=Column(ratio=1))
        bar_column = BarColumn(bar_width=None, table_column=Column(ratio=2))
        progress = Progress(text_column, bar_column, expand=True)
        console = Console()
        with console.status("[bold green]Parsing XML (2/3 minutes)...") :
            for reports in root:  # iterate over root
                for report in reports:
                    for content in progress.track(report):
                        for oval_results in content:
                            for results in oval_results:
                                for system in results:
                                    for oval_system_characteristics in system:
                                        if 'Group' in oval_system_characteristics.tag:
                                            recursive_iter_over_group(oval_system_characteristics, 0)
        """
        Third step is to create a csv file with all the info
        """

        nb_pt_ok = 0
        nb_pt_ko = 0
        nb_pt_manuel = 0
        nb_pt_na = 0

        mesures = []



        for entry in entry_list:

            #Synthese part
            mesure = {}

            current_result = entry.get_result()
            if current_result == "pass":
                nb_pt_ok = nb_pt_ok + 1
                mesure['statut'] = "OK"
                mesure['statut_bg'] = "2ECC71"
            elif current_result == "fail":
                nb_pt_ko = nb_pt_ko +1
                mesure['statut'] = "KO"
                mesure['statut_bg'] = "E74C3C"
            elif current_result == "manual":
                nb_pt_manuel = nb_pt_manuel + 1
                mesure['statut'] = "MANUEL"
                mesure['statut_bg'] = "BDC3C7"
            else:
                nb_pt_na = nb_pt_na + 1
                mesure['statut'] = "N/A"
                mesure['statut_bg'] = "Orange"

            #Mesure part

            mesure['id'] = entry.get_number()
            mesure['category'] = entry.get_branch()
            mesure['name'] = entry.get_control()
            mesure['description'] = entry.get_description()
            mesure['remediation'] = entry.get_remediation()
            mesures.append(mesure)

        synthese = {}
        synthese['nb_pt_ok'] = nb_pt_ok
        synthese['nb_pt_ko'] = nb_pt_ko
        synthese['nb_pt_manuel'] = nb_pt_manuel
        synthese['nb_pt_na'] = nb_pt_na

        create_report_cis(mesures, synthese, name)


        f = open("output", 'wb')

        f.write(bytes('Category, Title, Result, Remediation\n', 'UTF-8'))

        for entry in entry_list:
            f.write(bytes(entry.get_csv_string(), 'UTF-8'))

        f.close()

    else:
        logger.error("Currently working only on windows.")
        sys.exit(1)

    logger.info("End of parsing " + nameCisXmlFile + " file...")


def create_report_cis(mesures, synthese, name):
    logger.info("Creating Docx...")
    tpl = DocxTemplate('templates/template_word.docx')
    try:
        tpl.render(context={"synthese": synthese, "mesures": mesures})
    except Exception as e:
        logger.error(e)
        sys.exit(1)

    filename = unicode("results/" + name + ".docx")
    tpl.save(filename)
    logger.success("Docx successfully created at " + filename)

if __name__ == '__main__':
    options = parseArgs()
    logger.setVerbosity(options.verbose)

    for pathCisXmlFile in glob.iglob(".\\cis_xml\\*", recursive=False):
        parse_cis_html(pathCisXmlFile, options.export_docx)

