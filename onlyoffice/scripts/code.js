// @copyright Copyright (c) 2023 Nicolas GRELLETY
// @license https://opensource.org/licenses/GPL-3.0 GNU GPLv3
// @link https://gitlab.internal.lan/octo-project/octoaddins
// @link https://github.com/nillyr/octoaddins
// @since 0.1.0

var locales =
{
    "en-US": {
        "Synthesis": "Synthesis",
        "Conformity level": "Conformity level",
        "Success": "Success",
        "Failed": "Failed"
    },
    "fr-FR": {
        "Synthesis": "Synthèse",
        "Conformity level": "Niveau de conformité",
        "Success": "Succès",
        "Failed": "Échec"
    }
};

var lang = "en-US";
var selectLang = function(event) {
    var input = event.target;
    lang = input.value;
};

var fileContent = null;
var readFile = function(event) {
    var input = event.target;
    var reader = new FileReader();
    reader.onload = function() {
        var text = reader.result;
        fileContent = reader.result;
    };
    reader.readAsText(input.files[0]);
};

var getCSVContentFromText = function(data) {
    var options = {
        "separator": ",",
        "delimiter": '"'
    };
    return $.csv.toArrays(data, options);
};

var RuleResult = function(categoryReference, ruleReference, ruleName, ruleLevel, compliant) {
    this.categoryReference = categoryReference;
    this.ruleReference = ruleReference;
    this.ruleName = ruleName;
    this.ruleLevel = ruleLevel;
    this.compliant = compliant;
};

var capitalizeFirstLetter = function(string) {
    return string.charAt(0).toUpperCase() + string.slice(1);
}

var tableHeaders = [];
var convertToDict = function(data) {
    var isCSVHeader = true;
    var dict = {};
    for (var item of data) {
        if (isCSVHeader) {
            const regex = new RegExp(/(Category|Rule)\s|\s(de la règle)/, "gi");
            tableHeaders.push(capitalizeFirstLetter(item[2].replace(regex, "")));
            tableHeaders.push(capitalizeFirstLetter(item[4].replace(regex, "")));
            tableHeaders.push(capitalizeFirstLetter(item[3].replace(regex, "")));
            tableHeaders.push(capitalizeFirstLetter(item[6].replace(regex, "")));
            isCSVHeader = false;
            continue;
        }
        if (!dict[item[1]]) {
            dict[item[1]] = [];
        }
        dict[item[1]].push(new RuleResult(item[0], item[2], item[3], capitalizeFirstLetter(item[4]), item[6]));
    }
    return dict;
};

(function (window, undefined) {
    window.Asc.plugin.onTranslate = function() {
        document.getElementById("select_file").innerHTML = window.Asc.plugin.tr("Select a file");
        document.getElementById("select_lang").innerHTML = window.Asc.plugin.tr("Select report language");
        document.getElementById("en_radio_label").innerHTML = window.Asc.plugin.tr("English");
        document.getElementById("fr_radio_label").innerHTML = window.Asc.plugin.tr("French");
    };

    window.Asc.plugin.init = function() {
        switch(lang) {
            case "fr-FR":
            case "en-US":
                break;
            default:
                lang = "en-US";
        }
        this.resizeWindow(320, 150, 320, 150, 320, 150);
    };

    window.Asc.plugin.button = function(id) {
        // id = 0 = primary button = "Import results"
        if (id == 0) {
            var csvContent = getCSVContentFromText(fileContent);
            var results = convertToDict(csvContent);

            var nbSuccess = 0;
            var nbFailed = 0;
            for (var item of csvContent) {
                if (item[6]?.toLowerCase?.() === "true") {
                    nbSuccess++;
                } else if (item[6]?.toLowerCase?.() === "false") {
                    nbFailed++;
                }
            }

            // callCommand method is executed in its own context isolated from other js data
            // pass additional data to callCommand
            Asc.scope.st = [results, tableHeaders, nbSuccess, nbFailed, locales[lang]];
            this.callCommand(function() {
                var results = Asc.scope.st[0];
                var tableHeaders = Asc.scope.st[1];
                var nbSuccess = Asc.scope.st[2];
                var nbFailed = Asc.scope.st[3];
                var locales = Asc.scope.st[4];

                // Document
                var oDocument = Api.GetDocument();
                var oDocumentStyle = oDocument.GetStyle("Heading 2");
                var oParagraph = Api.CreateParagraph();
                oParagraph.SetStyle(oDocumentStyle);
                oParagraph.AddText("FIXME_ASSET_NAME");
                oParagraph.SetHighlight("yellow");
                oDocument.InsertContent([oParagraph]);

                oParagraph = Api.CreateParagraph();
                oDocumentStyle = oDocument.GetStyle("Heading 3");
                oParagraph.SetStyle(oDocumentStyle);
                oParagraph.AddText(locales["Synthesis"]);
                oDocument.InsertContent([oParagraph]);

                oParagraph = Api.CreateParagraph();
                oDocumentStyle = oDocument.GetStyle("Normal");
                oParagraph.SetStyle(oDocumentStyle);
                oParagraph.AddText("FIXME_SYNTHESIS");
                oParagraph.SetHighlight("yellow");
                oDocument.InsertContent([oParagraph]);

                oParagraph = Api.CreateParagraph();
                var series = [nbSuccess,nbFailed];
                // 1 mm = 36000 EMUs
                var nWidth = 150 * 36000; // 15 cm
                var nHeight = 100 * 36000; // 10 cm
                var nStyleIndex = 2; // "Office"
                var oChart = Api.CreateChart("pie", [
                    series
                ], ["percent"], [locales["Success"], locales["Failed"]], nWidth, nHeight, nStyleIndex);
                oChart.ApplyChartStyle(7);
                oChart.SetTitle(locales["Conformity level"], 13);
                oChart.SetLegendPos("right");
                oParagraph.AddDrawing(oChart);
                oParagraph.SetJc("center");
                oDocument.InsertContent([oParagraph]);

                // Create the style to use for every table
                var oTableStyle = oDocument.CreateStyle("octoconf", "table");
                oTableStyle.SetBasedOn(oDocument.GetStyle("Header"));
                var oTablePr = oTableStyle.GetTablePr();
                // Set all borders
                oTablePr.SetTableBorderTop("single", 4, 0, 0, 0, 0);
                oTablePr.SetTableBorderBottom("single", 4, 0, 0, 0, 0);
                oTablePr.SetTableBorderLeft("single", 4, 0, 0, 0, 0);
                oTablePr.SetTableBorderRight("single", 4, 0, 0, 0, 0);
                oTablePr.SetTableBorderInsideV("single", 4, 0, 0, 0, 0);
                oTablePr.SetTableBorderInsideH("single", 4, 0, 0, 0, 0);
                // Use full width
                oTablePr.SetWidth("percent", 100);
                // Align text
                oTablePr.SetJc("left");
                // Set margins
                oTablePr.SetTableCellMarginTop(28.346278133681); // 0.05cm
                oTablePr.SetTableCellMarginBottom(28.346278133681); // 0.05cm
                oTablePr.SetTableCellMarginLeft(107.7158569079878); // 0.19cm
                oTablePr.SetTableCellMarginRight(107.7158569079878); // 0.19cm
                // Enable auto resize
                oTablePr.SetTableLayout("autofit");
                // Set row height
                var oTableRowPr = oTableStyle.GetTableRowPr();
                oTableRowPr.SetHeight("atLeast", 340.15533760417196); // 0.6cm

                for (const [key, values] of Object.entries(results)) {
                    oParagraph = Api.CreateParagraph();
                    oDocumentStyle = oDocument.GetStyle("Heading 3");
                    oParagraph.SetStyle(oDocumentStyle);
                    oParagraph.AddText(key);
                    oDocument.InsertContent([oParagraph]);

                    var oTable = Api.CreateTable(4, 1);
                    oTable.SetStyle(oTableStyle);
                    oTable.SetTableLook(true, true, true, true, true, true);
                    var oTableRow = oTable.GetRow(0);
                    oTableRow.SetBackgroundColor(51,62,78,false);
                    oTableRow.SetTableHeader(true);

                    var colIndex = 0;
                    for (var header of tableHeaders) {
                        var oCell = oTableRow.GetCell(colIndex);
                        var oDocumentElt = oCell.GetContent().GetElement(0);
                        oDocumentElt.AddText(header);
                        oDocumentElt.SetBold(true);
                        oDocumentElt.SetColor(255,255,255);
                        oTable.AddElement(oCell, 0, oDocumentElt);
                        colIndex++;
                    }
                    // Commit
                    oDocument.InsertContent([oTable]);

                    var rowIndex = 1;
                    for (var value of values) {
                        oTable.AddRow(null, true);
                        var oTableRow = oTable.GetRow(rowIndex);
                        oTableRow.SetBackgroundColor(255,255,255,false);
                        oTableRow.SetTableHeader(false);

                        var cellsContent = [
                            value["ruleReference"],
                            value["ruleLevel"],
                            value["ruleName"],
                            value["compliant"]
                        ];

                        var colIndex = 0;
                        for (var cellContent of cellsContent) {
                            var oCell = oTableRow.GetCell(colIndex);
                            var oDocumentElt = oCell.GetContent().GetElement(0);
                            oDocumentElt.AddText(cellContent);
                            switch(cellContent?.toLowerCase?.()) {
                                case 'minimal':
                                    oDocumentElt.SetBold(true);
                                    oDocumentElt.SetColor(197,23,24,false);
                                    break;
                                case 'intermediary':
                                    oDocumentElt.SetBold(true);
                                    oDocumentElt.SetColor(241,153,45,false);
                                    break;
                                case 'enhanced':
                                    oDocumentElt.SetBold(true);
                                    oDocumentElt.SetColor(255,204,0,false);
                                    break;
                                case 'high':
                                    oDocumentElt.SetBold(true);
                                    oDocumentElt.SetColor(0,150,68,false);
                                    break;
                                case 'true':
                                    oDocumentElt.SetBold(true);
                                    oDocumentElt.SetColor(0,150,68,false);
                                    break;
                                case 'false':
                                    oDocumentElt.SetBold(true);
                                    oDocumentElt.SetColor(197,23,24,false);
                                    break;
                                default:
                                    oDocumentElt.SetBold(false);
                                    oDocumentElt.SetColor(0,0,0,false);
                            }
                            oTable.AddElement(oCell, 0, oDocumentElt);
                            colIndex++;
                        }
                        // Commit
                        oDocument.InsertContent([oTable]);
                        rowIndex++;
                    }

                    oParagraph = Api.CreateParagraph();
                    oDocumentStyle = oDocument.GetStyle("Heading 4");
                    oParagraph.SetStyle(oDocumentStyle);
                    oParagraph.AddText(locales["Synthesis"]);
                    oDocument.InsertContent([oParagraph]);

                    oParagraph = Api.CreateParagraph();
                    oParagraph.AddText("FIXME_CATEGORY_SYNTHESIS");
                    oParagraph.SetHighlight("yellow");
                    oDocument.InsertContent([oParagraph]);
                }
            }, true);
        } else {
            this.executeCommand("close", "");
        }
    };
})(window, undefined);
