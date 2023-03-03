#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
#  ODFTemplator.py
#  PythonODFTemplator
#
#  Copyright 2013-2023 Kenji Nishishiro. All rights reserved.
#  Written by Kenji Nishishiro <marvel@programmershigh.org>.
#

import os
import re
import subprocess
import tempfile
import zipfile

import jinja2
import pylokit


class TemplatingError(Exception):
    pass


class Templator(object):
    def __init__(self, libreoffice_method="kit", libreoffice_path="/usr/lib/libreoffice/program"):
        self.environment = jinja2.Environment(autoescape=True)

        self.__libreoffice_method = libreoffice_method
        self.__libreoffice_path = libreoffice_path

    def render(self, template_file_path, document_file_path, **params):
        try:
            with zipfile.ZipFile(template_file_path, "r") as source_file:
                with zipfile.ZipFile(document_file_path, "w") as destination_file:
                    for name in source_file.namelist():
                        data = source_file.read(name)
                        if name == "content.xml":
                            template = self.environment.from_string(self.fix_block(data.decode("utf-8")))
                            data = template.render(**params).encode("utf-8")
                        destination_file.writestr(name, data)
        except IOError as error:
            raise TemplatingError(*error.args) from error
        except jinja2.TemplateError as error:
            raise TemplatingError(*error.args) from error

    def render_pdf(self, template_file_path, pdf_file_path, **params):
        try:
            with tempfile.TemporaryDirectory() as document_directory_path:
                document_file_path = os.path.join(
                    document_directory_path,
                    os.path.splitext(os.path.basename(pdf_file_path))[0] + os.path.splitext(template_file_path)[1],
                )
                self.render(template_file_path, document_file_path, **params)

                if self.__libreoffice_method == "kit":
                    with pylokit.Office(self.__libreoffice_path) as office:
                        with office.documentLoad(document_file_path) as document:
                            document.saveAs(pdf_file_path)
                elif self.__libreoffice_method == "command":
                    subprocess.call(
                        [
                            self.__libreoffice_path,
                            "--convert-to",
                            "pdf",
                            "--outdir",
                            os.path.dirname(pdf_file_path),
                            document_file_path,
                        ]
                    )
                    # TODO: Check failure reason
                    if not os.path.isfile(pdf_file_path):
                        raise TemplatingError("Conversion error occurred.")
                else:
                    raise TemplatingError("Unsupported libreoffice method.")
        except IOError as error:
            raise TemplatingError(*error.args) from error
        except pylokit.LoKitInitializeError as error:
            raise TemplatingError(*error.args) from error
        except pylokit.LoKitImportError as error:
            raise TemplatingError(*error.args) from error
        except pylokit.LoKitExportError as error:
            raise TemplatingError(*error.args) from error

    @staticmethod
    def fix_block(content):
        def repl(match):
            text = match.group(0)
            text = text.replace("&quot;", '"')
            text = text.replace("&amp;", "&")
            text = text.replace("&lt;", "<")
            text = text.replace("&gt;", ">")
            text = text.replace("&apos;", "'")
            return text

        return re.sub(r"({[{%].+?[%}]})", repl, content)
