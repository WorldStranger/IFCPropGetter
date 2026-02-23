#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""程序主逻辑入口"""

from .gui import IFCPropertyExtractorApp

def main():
    app = IFCPropertyExtractorApp()
    app.run()