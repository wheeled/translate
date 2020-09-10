#!/usr/bin/env python

from __future__ import absolute_import

# TODO: look at multi-threading to speed up translation (probably can't do that with text file)

import argparse
from importlib import import_module
import logging
import pprint
import os
import sys
from google.cloud.translate import Client

from translate.translate_base import CREDS, GoogleTranslate

if not sys.warnoptions:
    import warnings
    warnings.simplefilter('default')

pp = pprint.PrettyPrinter(indent=4)


class ListLanguages(object):
    def __init__(self):
        self.client = Client.from_service_account_json(CREDS)
        self.print_response(self.request_languages())

    def request_languages(self):
        """ Lists all available languages. """
        # translate_client = Client.from_service_account_json(CREDS)
        return self.client.get_languages()

    def print_response(self, results):
        for language in results:
            print(u'{name} ({language})'.format(**language))


def parse_args(args, info):
    parser = argparse.ArgumentParser(
        prog='translate',
        description='Script to translate files from one language to another using Google Cloud Translate '
                    'API.  Formats supported: %s.' % info,
        epilog='Requires credentials for Google Cloud to be saved.'
    )
    parser.add_argument('file', help='file to be translated')
    parser.add_argument('target_lang', help='target language per ISO 639-1')
    # parser.add_argument('-a', '--auth', dest='creds', help='Google Cloud API credentials file')
    parser.add_argument('-c', '--condense', dest='condense', default=False, action='store_true',
                        help='condense runs in paragraph')
    parser.add_argument('-d', '--dest', dest='target', help='translation output filename')
    # TODO: prefer to make -l behave like -h so it does not require the positional arguments
    parser.add_argument('-l', '--list_langs', default=False, action='store_true',
                        help='print list of all available languages')
    # parser.add_argument('-m', '--threads', dest='threads', type=int, default='1',
    #                     help='use N threads to access cloud')
    parser.add_argument('-n', '--preview', dest='online', default=True, action='store_false',
                        help='offline preview mode')
    parser.add_argument('-p', '--progress', dest='show', type=int, default='0', help='show N chars of each string')
    parser.add_argument('-q', '--quiet', action='count', default=0, help='decrease logging level')
    parser.add_argument('-r', '--reuse', dest='history', help='filename (for reuse of translation strings)')
    parser.add_argument('-s', '--source', dest='source_lang', help='source language per ISO 639-1')
    parser.add_argument('-v', '--verbose', action='count', default=0, help='increase logging level')
    parser.add_argument('-x', '--xcheck', dest='cross_check', default=False, action='store_true',
                        help='translate back again for checking')
    return parser.parse_args(args)


def main(arg_list=None):
    TRANSLATOR = {
        'docx': 'translate.translate_base.TranslateDocx',
        'pptx': 'translate.translate_base.TranslatePptx',
        'txt': 'translate.translate_base.TranslateText',
        'xlsx': 'translate.translate_base.TranslateExcel',
    }

    logging.basicConfig(level=logging.WARNING)
    log = logging.getLogger(__name__)
    args = parse_args(arg_list, ', '.join(sorted(TRANSLATOR.keys())))
    if os.path.isfile(os.path.realpath(args.file)):
        args.filepath, args.filename = os.path.split(os.path.realpath(args.file))
    else:
        # TODO: handle error better
        warnings.warn("Not a valid path to a file '%s'" % args.file)
        return "invalid file"

    log.setLevel(max(1, logging.WARNING + (args.quiet - args.verbose) * 10))

    # placeholder for creds changes

    if args.list_langs:
        ListLanguages()

    else:
        # placeholder for multi-threading - need to limit value of threads

        bf_kwargs = {key: value for (key, value) in vars(args).items()
                     if key in ['source_lang', 'online', 'history', 'show']}
        babelfish = GoogleTranslate(CREDS, args.target_lang, **bf_kwargs)

        ft_kwargs = {key: value for (key, value) in vars(args).items()
                     if key in ['condense', 'cross_check']}

        file_format = args.filename.rsplit('.', 1)[1]
        module_name, class_name = TRANSLATOR[file_format].rsplit('.', 1)
        file_translator = getattr(import_module(module_name), class_name)
        file_translator(args.filepath, args.filename, babelfish, **ft_kwargs)

        if hasattr(babelfish, 'stats'):
            for lang_pair in babelfish.stats:
                print('statistics for %s translation session', lang_pair)
                for stat in babelfish.stats[lang_pair]:
                    print('    %s: %s' % (stat, babelfish.stats[lang_pair][stat]))


def init():
    if __name__ == '__main__':
        sys.exit(main())


init()
