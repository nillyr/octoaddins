# octoconf-addin

## Installation process

See this [documentation](https://api.onlyoffice.com/plugin/installation/desktop).

## Contributing

See [A quick start guide for developers](https://www.onlyoffice.com/blog/2020/04/plugins-in-onlyoffice-a-quick-start-guide-for-developers). You can find the full API documentation [here](https://api.onlyoffice.com/docbuilder/textdocumentapi).

On GNU/Linux, you may want to add a new Desktop entry with the following Exec value:

```bash
# Change type with either "word", "cell" or "slide"
Exec=/usr/bin/onlyoffice-desktopeditors --ascdesktop-support-debug-info --new:<type>
```

See [Running ONLYOFFICE Desktop Editors with parameters](https://helpcenter.onlyoffice.com/installation/desktop-flags.aspx).

When running the addin, right-click on the `iframe` and click on `Show DevTools`.

## Maintainer

- Nicolas GRELLETY

## Authors

- Nicolas GRELLETY

## Copyright and license

Copyright (c) 2023 Nicolas GRELLETY

This software is licensed under GNU GPLv3 license. See `LICENSE` file in the root folder of the project.

Icons: [Import CSV](https://icons8.com/icon/32516/import-csv) icon by [Icons8](https://icons8.com/)
