// @copyright Copyright (c) 2023 Nicolas GRELLETY
// @license https://opensource.org/licenses/GPL-3.0 GNU GPLv3
// @link https://gitlab.internal.lan/octo-project/octoaddins
// @link https://github.com/nillyr/octoaddins
// @since 0.1.0

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

var RuleResult = function(category_reference, rule_reference, rule_name, rule_level, rule_severity, compliant) {
    this.category_reference = category_reference;
    this.rule_reference = rule_reference;
    this.rule_name = rule_name;
    this.rule_level = rule_level;
    this.rule_severity = rule_severity;
    this.compliant = compliant;
};

var convertToDict = function(data) {
    var isCSVHeader = true;
    var dict = {};
    for (var item of data) {
        if (isCSVHeader) {
            isCSVHeader = false;
            continue;
        }
        if (!dict[item[1]]) {
            dict[item[1]] = [];
        }
        dict[item[1]].push(new RuleResult(item[0], item[2], item[3], item[4], item[5], item[6]));
    }
    return dict;
};

(function (window, undefined) {
    window.Asc.plugin.init = function () {
    };

    window.Asc.plugin.button = function (id) {
        // id = 0 = primary button = "Import results"
        if (id == 0) {
            var csv_content = getCSVContentFromText(fileContent);
            var results_dict = convertToDict(csv_content);

            var nbSuccess = 0;
            var nbFailed = 0;
            for (var item of csv_content) {
                if (item[6]?.toLowerCase?.() === "true") {
                    nbSuccess++;
                } else if (item[6]?.toLowerCase?.() === "false") {
                    nbFailed++;
                }
            }

            // callCommand method is executed in its own context isolated from other js data
            // pass additional data to callCommand
            Asc.scope.st = [results_dict, nbSuccess, nbFailed];
            this.callCommand(function() {
                var results_dict = Asc.scope.st[0];
                var nbSuccess = Asc.scope.st[1];
                var nbFailed = Asc.scope.st[2];

                var oDocument = Api.GetDocument();
                var oDocumentStyle = oDocument.GetStyle("Heading 2");
                var oParagraph = Api.CreateParagraph();
                oParagraph.SetStyle(oDocumentStyle);
                oParagraph.AddText("FIXME_ASSET_NAME");
                oDocument.InsertContent([oParagraph]);

                oParagraph = Api.CreateParagraph();
                oDocumentStyle = oDocument.GetStyle("Heading 3");
                oParagraph.SetStyle(oDocumentStyle);
                oParagraph.AddText("Synthesis");
                oDocument.InsertContent([oParagraph]);

                oParagraph = Api.CreateParagraph();
                oDocumentStyle = oDocument.GetStyle("Normal");
                oParagraph.SetStyle(oDocumentStyle);
                oParagraph.AddText("FIXME: add your global synthesis based on the results");
                oDocument.InsertContent([oParagraph]);

                oParagraph = Api.CreateParagraph();
                var series = [nbSuccess,nbFailed];
                // 1 mm = 36000 EMUs
                var nWidth = 150 * 36000; // 15 cm
                var nHeight = 100 * 36000; // 10 cm
                var nStyleIndex = 2; // "Office"
                var oChart = Api.CreateChart("pie", [
                    series
                ], ["percent"], ["success", "failed"], nWidth, nHeight, nStyleIndex);
                oChart.ApplyChartStyle(7);
                oChart.SetTitle("Conformity level", 13);
                oChart.SetLegendPos("right");
                oParagraph.AddDrawing(oChart);
                oParagraph.SetJc("center");
                oDocument.InsertContent([oParagraph]);

                // FIXME: ajout de la possiblité de choisir la couleur / font style du header du tableau
                for (const [key, values] of Object.entries(results_dict)) {
                    oParagraph = Api.CreateParagraph();
                    oDocumentStyle = oDocument.GetStyle("Heading 3");
                    oParagraph.SetStyle(oDocumentStyle);
                    oParagraph.AddText(key);
                    oDocument.InsertContent([oParagraph]);

                    // +1 for the header
                    var oTable = Api.CreateTable(5, values.length + 1);
                    oTable.SetWidth("percent", 100);

                    oParagraph = Api.CreateParagraph();
                    oParagraph.AddText("Reference");
                    var oCell = oTable.GetCell(0,0);
                    oTable.AddElement(oCell, 0, oParagraph);

                    oParagraph = Api.CreateParagraph();
                    oParagraph.AddText("Level");
                    oCell = oTable.GetCell(0,1);
                    oTable.AddElement(oCell, 0, oParagraph);

                    oParagraph = Api.CreateParagraph();
                    oParagraph.AddText("Name");
                    oCell = oTable.GetCell(0,2);
                    oTable.AddElement(oCell, 0, oParagraph);

                    oParagraph = Api.CreateParagraph();
                    oParagraph.AddText("Severity");
                    oCell = oTable.GetCell(0,3);
                    oTable.AddElement(oCell, 0, oParagraph);

                    oParagraph = Api.CreateParagraph();
                    oParagraph.AddText("Status");
                    oCell = oTable.GetCell(0,4);
                    oTable.AddElement(oCell, 0, oParagraph);
                    // commit
                    oDocument.InsertContent([oTable]);

                    var row_index = 1;
                    for (var value of values) {
                        oParagraph = Api.CreateParagraph();
                        oParagraph.AddText(value["rule_reference"]);
                        oCell = oTable.GetCell(row_index,0);
                        oTable.AddElement(oCell, 0, oParagraph);

                        oParagraph = Api.CreateParagraph();
                        oParagraph.AddText(value["rule_level"]);
                        oCell = oTable.GetCell(row_index,1);
                        oTable.AddElement(oCell, 0, oParagraph);

                        oParagraph = Api.CreateParagraph();
                        oParagraph.AddText(value["rule_name"]);
                        oCell = oTable.GetCell(row_index,2);
                        oTable.AddElement(oCell, 0, oParagraph);

                        oParagraph = Api.CreateParagraph();
                        oParagraph.AddText(value["rule_severity"]);
                        oCell = oTable.GetCell(row_index,3);
                        oTable.AddElement(oCell, 0, oParagraph);

                        oParagraph = Api.CreateParagraph();
                        oParagraph.AddText(value["compliant"]);
                        oCell = oTable.GetCell(row_index,4);
                        oTable.AddElement(oCell, 0, oParagraph);
                        // commit
                        oDocument.InsertContent([oTable]);

                        row_index++;
                    }

                    oParagraph = Api.CreateParagraph();
                    oDocumentStyle = oDocument.GetStyle("Heading 4");
                    oParagraph.SetStyle(oDocumentStyle);
                    oParagraph.AddText("Synthesis");
                    oDocument.InsertContent([oParagraph]);

                    oParagraph = Api.CreateParagraph();
                    oDocumentStyle = oDocument.GetStyle("Normal");
                    oParagraph.SetStyle(oDocumentStyle);
                    oParagraph.AddText("FIXME: add your synthesis based on the results");
                    oDocument.InsertContent([oParagraph]);
                }
            }, true);
            // [0] = Référence de la catégorie,Nom de la catégorie,Référence de la règle,Nom de la règle,Niveau de la règle,Sévérité de la règle,Conforme
            // [1] = hardware_support,Support matériel,NP_ConfigMateriel_R1,Support des instructions 64 bits par le processeur,enhanced,unknown,True

            // https://api.onlyoffice.com/docbuilder/textdocumentapi/api/getdocument

            // https://api.onlyoffice.com/docbuilder/textdocumentapi/apiparagraph
            // https://api.onlyoffice.com/docbuilder/textdocumentapi/apiparagraph/addheadingcrossref <<========HERE !!!! :)
            // https://api.onlyoffice.com/docbuilder/textdocumentapi/apiparagraph/addlinebreak

            // https://api.onlyoffice.com/docbuilder/textdocumentapi/api/createchart
            // https://api.onlyoffice.com/docbuilder/global#ChartType

            // https://api.onlyoffice.com/docbuilder/textdocumentapi/apitable
            // https://api.onlyoffice.com/docbuilder/textdocumentapi/api/createtable
            // https://api.onlyoffice.com/docbuilder/textdocumentapi/apitable/addelement
        } else {
            this.executeCommand("close", "");
        }
    };
})(window, undefined);
