# plugin-octoconf

https://www.onlyoffice.com/blog/2020/04/plugins-in-onlyoffice-a-quick-start-guide-for-developers

https://api.onlyoffice.com/plugin/example/helloworld

https://api.onlyoffice.com/plugin/installation/desktop

asc.{57E213AA-0804-44C5-92D8-769B1E1BFD42}

https://api.onlyoffice.com/plugin/basic

https://dzone.com/articles/extending-the-onlyoffice-online-editors-functional

choise: download local script/css so this plugin can be installed on restricted env (no internet)

https://www.tecmint.com/create-plugin-onlyoffice-docs/

```bash
cd /opt/onlyoffice/desktopeditors/editors/sdkjs-plugins/\{57E213AA-0804-44C5-92D8-769B1E1BFD42\}
rm -rf *; cp -r /home/ngy/Documents/Git/octo\ project/octoaddins/onlyoffice/* .
```

Icons: [Import CSV](https://icons8.com/icon/32516/import-csv) icon by [Icons8](https://icons8.com/)

Add debug support `--ascdesktop-support-debug-info`:

```ini
[Desktop Entry]
Version=1.0
Name=(DEV) - ONLYOFFICE Desktop Editors
GenericName=(DEV) - Document Editor
GenericName[ru]=Редактор документов
Comment=(DEV) - Edit office documents
Comment[ru]=Редактировать офисные документы
Type=Application
Exec=/usr/bin/onlyoffice-desktopeditors --ascdesktop-support-debug-info %U
Terminal=false
Icon=onlyoffice-desktopeditors
Keywords=Text;Document;OpenDocument Text;Microsoft Word;Microsoft Works;odt;doc;docx;rtf;
Categories=Office;WordProcessor;Spreadsheet;Presentation;
MimeType=application/vnd.oasis.opendocument.text;application/vnd.oasis.opendocument.text-template;application/vnd.oasis.opendocument.text-web;application/vnd.oasis.opendocument.text-master;application/vnd.sun.xml.writer;application/vnd.sun.xml.writer.template;application/vnd.sun.xml.writer.global;application/msword;application/vnd.ms-word;application/x-doc;application/rtf;text/rtf;application/vnd.wordperfect;application/wordperfect;application/vnd.openxmlformats-officedocument.wordprocessingml.document;application/vnd.ms-word.document.macroenabled.12;application/vnd.openxmlformats-officedocument.wordprocessingml.template;application/vnd.ms-word.template.macroenabled.12;application/vnd.oasis.opendocument.spreadsheet;application/vnd.oasis.opendocument.spreadsheet-template;application/vnd.sun.xml.calc;application/vnd.sun.xml.calc.template;application/msexcel;application/vnd.ms-excel;application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;application/vnd.ms-excel.sheet.macroenabled.12;application/vnd.openxmlformats-officedocument.spreadsheetml.template;application/vnd.ms-excel.template.macroenabled.12;application/vnd.ms-excel.sheet.binary.macroenabled.12;text/csv;text/spreadsheet;application/csv;application/excel;application/x-excel;application/x-msexcel;application/x-ms-excel;text/comma-separated-values;text/tab-separated-values;text/x-comma-separated-values;text/x-csv;application/vnd.oasis.opendocument.presentation;application/vnd.oasis.opendocument.presentation-template;application/vnd.sun.xml.impress;application/vnd.sun.xml.impress.template;application/mspowerpoint;application/vnd.ms-powerpoint;application/vnd.openxmlformats-officedocument.presentationml.presentation;application/vnd.ms-powerpoint.presentation.macroenabled.12;application/vnd.openxmlformats-officedocument.presentationml.template;application/vnd.ms-powerpoint.template.macroenabled.12;application/vnd.openxmlformats-officedocument.presentationml.slide;application/vnd.openxmlformats-officedocument.presentationml.slideshow;application/vnd.ms-powerpoint.slideshow.macroEnabled.12;x-scheme-handler/oo-office;text/docxf;text/oform;
Actions=NewDocument;NewSpreadsheet;NewPresentation;

[Desktop Action NewDocument]
Name=New Document
Name[de]=Neues Dokument
Name[fr]=Nouveau document
Name[es]=Documento nuevo
Name[ru]=Создать документ
Exec=/usr/bin/onlyoffice-desktopeditors --ascdesktop-support-debug-info --new:word

[Desktop Action NewSpreadsheet]
Name=New Spreadsheet
Name[de]=Neues Tabellendokument
Name[fr]=Nouveau classeur
Name[es]=Hoja de cálculo nueva
Name[ru]=Создать эл.таблицу
Exec=/usr/bin/onlyoffice-desktopeditors --ascdesktop-support-debug-info --new:cell

[Desktop Action NewPresentation]
Name=New Presentation
Name[de]=Neue Präsentation
Name[fr]=Nouvelle présentation
Name[es]=Presentación nueva
Name[ru]=Создать презентацию
Exec=/usr/bin/onlyoffice-desktopeditors --ascdesktop-support-debug-info --new:slide
```
