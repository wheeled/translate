``translate``
=============

Overview
---------
``translate`` is a Python utility to use Google Cloud Translate API to translate documents from one language to another.

The script does the heavy lifting by using other open source libraries, in particular:

* google-cloud-translate https://cloud.google.com/translate/docs/quickstart
* openpyxl https://openpyxl.readthedocs.io/en/stable/
* python-docx https://python-docx.readthedocs.io/en/latest/
* python-pptx https://python-pptx.readthedocs.io/en/latest/

The objective of ``translate`` is to "do no harm" by treading lightly over an existing document, finding instances of text that can be translated, and replacing that text in the original document while disturbing as little else as possible in the process.  This contrasts with the philosophy of the above three Open XML projects, which is more about programmatically implementing the Office Open XML standards (https://www.ecma-international.org/publications/standards/Ecma-376.htm) and to read and/or compose documents in compliance with the standard.

``translate`` assumes that many, if not most, of the documents being translated will have been originally created by another application, such as Microsoft Word.  As a result there is more focus on reverse engineering to work around the quirks of the source applications than might otherwise be the case.

Important Note
--------------
Any text you submit to Google is not secure.  You should not use this tool to translate confidential information.

Installation
------------
The ``translate`` repo is not presently published so that it can not be found by ``pip``, ``brew`` or ``easy-install``.  Installation involves cloning the repo onto your machine and then installing from there::

    hg clone https://WheeleD@bitbucket.org/WheeleD/translate
    pip3 install -e /path/to/your/local/repo

Usage
-----
First and foremost, to leverage the Google Cloud Translate API, Google Cloud service credentials are required.  The terms on which Google (presently) allows access to this API are very generous, and the process of obtaining and installing credentials is explained at the quickstart link above.

Once installed, ``translate`` can be called from the command line with as little as the name of the file to be translated and the language it needs to be translated into.

Example::

  $ python3 path/to/translate my_document.docx fr

There are also python classes which can be imported into your own scripts to perform translation as an element of a larger body of work.  There are three types:

GoogleTranslate
    This class uses the credentials to establish an http session with Google Cloud, and provides a ``translate`` method which returns the translated version of a single string.

TranslateBase
    This class implements the base functionality for a file-format-specific translation class, and allows extensibility to other file formats by developing subclasses that support specific file formats.

TranslateXxxx
    A subclass of ``TranslateBase``, these classes deal with the specific nature of a specific file format to iterate over the text instances in the file and deliver them to an instance ``GoogleTranslate`` class before finally replacing them in the file.

The usage in a script is also very straightforward::

        from translate import GoogleTranslate, TranslateDocx

        babelfish = GoogleTranslate('fr')
        TranslateDocx('/users/me/documents', 'my_document.docx', babelfish)

Note the above example requires installation, and the existence of credentials as an environment variable. The instance name in the example, ``babelfish``, is a nod to Douglas Adams who foresaw this capability.

The command line help provides an overview of the command line argument options::

    usage: translate [-h] [-c] [-d TARGET] [-l] [-n] [-p SHOW] [-q] [-r HISTORY]
                     [-s SOURCE_LANG] [-v] [-x]
                     file target_lang

    Script to translate files from one language to another using Google Cloud
    Translate API. Formats supported: docx, pptx, txt, xlsx.

    positional arguments:
      file                  file to be translated
      target_lang           target language per ISO 639-1

    optional arguments:
      -h, --help            show this help message and exit
      -c, --condense        condense runs in paragraph
      -d TARGET, --dest TARGET
                            translation output filename
      -l, --list_langs      print list of all available languages
      -n, --preview         offline preview mode
      -p SHOW, --progress SHOW
                            show N chars of each string
      -q, --quiet           decrease logging level
      -r HISTORY, --reuse HISTORY
                            filename (for reuse of translation strings)
      -s SOURCE_LANG, --source SOURCE_LANG
                            source language per ISO 639-1
      -v, --verbose         increase logging level
      -x, --xcheck          translate back again for checking

    Requires credentials for Google Cloud to be saved.

The class definitions show the kwargs when instantiating the classes from a script::

    class GoogleTranslate(object):
        """ Establish a Google Cloud Translate client to translate passages of text. """
        def __init__(self, creds, target_lang, source_lang=None, online=True, history=None, show=0):
            ...

    class TranslateText(TranslateBase):
        """ Translate text in a Word (.docx) document file """
        def __init__(self, filepath, filename, translator, target=None, condense=False, cross_check=False):
            ...

The kwargs for the file-format-specific classes are passed using super to ``TranslateBase``.

A brief explanation of the arguments for these classes follows.

creds
    the credentials required for access to Google Cloud, obtained either from the environment variable or the path to the JSON credentials file provided by Google.

target_lang
    a two-letter string identifying the language required, according to the ISO 639-1 standard (second column at https://www.loc.gov/standards/iso639-2/php/code_list.php).

source_lang (optional)
    default None.  Can be supplied as a two-letter ISO 639-1 code.  It is not required because Google Cloud Translate will auto-detect if it is not supplied.  If a single source language is present in the source file it is advisable to be specific or some unintended translations are possible (notably of acronyms).

online (optional)
    default ``True``.  If ``False``, will mark the boundaries of each text string that would be submitted for translation, without actually translating the text.  This can be helpful for understanding patchy translation results (that can occur due to the way Microsoft Word marks edits in a *docx* file.)

history (optional)
    default None.  If a filepath is specified here, then the translation dictionary can be saved on completion of translation, and will be loaded and used to minimise the number of API calls for future translations of the same file.  If an argument is passed that does not resolve into a filename, the history file will be saved as the body of ``filename`` with an underscore and the ``target_lang`` code, for example ``my_document_fr.json``.

show (optional)
    default 0.  If set to a positive integer N, will show the first N letters of each string submitted for translation.  This serves as a progress indicator as well as helping to identify truncated words that are being sent and therefore result in mis-translation.

filepath
    the full path to the directory containing the source file.  If called from the command line, this will be derived from the ``file`` argument.

filename
    the name of the source file.  If called from the command line, this will be derived from the ``file`` argument.

translator
    an instance of the ``GoogleTranslate`` class.

target (optional)
    the filename to use for the translated output file.  If not provided, the default is to use ``filepath`` as the location and extend the body of ``filename`` with an underscore and the ``target_lang`` code, for example ``my_document_fr.docx``.

condense (optional)
    default ``False``.  Use the ``-c`` argument or pass ``True`` to allow successive text runs in a paragraph to be concatenated into a single run for translation.  Styling changes or line feeds within the paragraph will terminate any concatenation.  This is desirable to deal with a quirk of Word and Powerpoint, whereby incomplete runs are frequently created where corrections are made.

cross-check (optional)
    default ``False``.  Use the ``-x`` argument or pass ``True`` to translate the document back into the source language for review.  This is a sanity check only, but can serve as an indication of the quality of the original translation.  Mistakes that show up here could suggest other ways of writing or laying out the original document to improve the translation.  No filename can be specified for this - the output will be saved as the body of ``target`` filename with an underscore and the ``source_lang`` code, for example ``my_document_fr_en.docx``.

Capabilities
------------
The ``translate`` tool is heavily dependent on the Office Open XML libraries, and the scope of their coverage of documents created by commercial applications.  The following notes, while not comprehensive, illustrate what can and can't be achieved.  Some of the unsupported capabilities would be possible by employing unreleased branches of the library - with one exception, ``translate`` relies on the latest released branch for stability and ease of installation.

DOCX format

* Translates all of the following:

 - Headings, preserving numbering and bookmarks for cross-reference and tables of contents
 - Paragraphs of body text, preserving any styling including in-line emphasis or font attributes and preserving embedded images and line feeds
 - Tables, preserving table styling and cell contents as for body text paragraphs
 - Captions of figures and tables, preserving numbering and bookmarks for cross-reference and tables of contents
 - (indirectly) tables of contents, figures and tables - these need to be manually updated [right click, Update field, Update entire table] when the translated document is opened in Word
 - (updates) Language attribute in Document Properties
 - (updates) Normal style Font (for Japanese)

* Does not translate:

 - Document property attribute values
 - Headers and Footers
 - Text boxes
 - (does not update) paragraph-level and below language attributes used by spell-checker

XLSX format

* Translates the following:

 - Cell text, preserving formulas and cell-level styling (with the exception of merged cell borders in OpenPyxl 2.5.5 - this is coming)

* Does not translate:

 - Headers and Footers
 - Worksheet tab names
 - (does not preserve) Rich text formatting within the cell
 - (does not preserve) Embedded images (this is coming in OpenPyxl)
 - (does not preserve) Column width settings for tables - after translation, table will have all columns as autofit.

PPTX format

* Translates all of the following:

 - Slide Titles and Bullets, preserving any styling including in-line emphasis or font attributes and preserving embedded line feeds
 - Tables, preserving any styling including in-line emphasis or font attributes and preserving embedded line feeds
 - Text Boxes and Shapes (apart from text in SmartArt shapes)
 - (updates) Language attribute in Document Properties - although this does not appear to be used by PowerPoint

* Does not translate:

 - Document property attribute values
 - Speaker's notes

Other Notes
-----------
There is a conflict between preserving formatting within a paragraph and achieving the best translation.  In this iteration of ``translate`` the assumption is that the translation achieved using this tool is a good start, but will require polishing by a native speaker to produce the finished product.  In that case, the preservation of the formatting serves as an indication to the native speaker that the author intended to emphasize certain points and she can take this into account.

If the tool will be producing the finished product (as a 'good enough' translation for the job at hand) then it would be better to sacrifice the formating within the paragraph so that the entire paragraph can be submitted to Google at once.  Looking for feedback on this point - disabling this feature could be implemented as an additional keyword argument.

03-Sep-2018
