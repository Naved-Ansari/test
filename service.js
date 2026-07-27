/*
Author: ANSARNA
CreatedOn: 24 July, 2026
LastEdited: 27 July, 2026
ChangeLog: Performance optimization - replaced O(n²) InfoTable queries with hash-map lookups
*/

try 
{
    var result, reA, reN, aA, bA, aN, bN, params;
    var z, x, i, j, k, sort, row, row1, row3, keyC;
    var refArray, refDesForTPQ = "0";
    var sourceRow;
    var newMasterEntry, newEntry, newEntry1;
    var overpopulatedBOM, overpopulatedBOMLength;
    var consolidatedBomRow, consolidatedBom;
    var pptaResponseTable, filteredPptaResponseTable, pptaResponseJSON;
    var selectedVariantCountTable;
    var allBomsInSingleTable;
    var allConsolidatedBomsTableLength;
    var variantTableLength;
    var totalPopulationQtyTable;
    var selectedVariantTable, tempTable, tempResult, masterBom, variantRow, allBomsRow;
    var allAxlEntriesTable, axlTable, axlSamePartExhistTable;
    var supplierCount = 0;
    var cCRTLString;
    var ecadcCRTLTable, ecadcCRTLRow;
    var refDes, part1, part2;

    //function to sort alphanumeric array
    reA = /[^a-zA-Z]/g;
    reN = /[^0-9]/g;

    function sortAlphaNum(a, b) 
    {
        aA = a.replace(reA, "");
        bA = b.replace(reA, "");
        if (aA === bA) 
        {
            aN = parseInt(a.replace(reN, ""), 10);
            bN = parseInt(b.replace(reN, ""), 10);
            return aN === bN ? 0 : aN > bN ? 1 : -1;
        } 
        else 
        {
            return aA > bA ? 1 : -1;
        }
    }

    params = {
        infoTableName: "cc" /* STRING */
    };
    finalTable = Resources["InfoTableFunctions"].CreateInfoTable(params);


    if (masterOid !== undefined && masterOid !== "") 
    {
        try 
        {
            //get Master Bom
            logger.info("PrototypeBOM get master BOM start");
            masterBom = me.getBomInfotableForPrototypeBom({
                WTPartOid: masterOid /* STRING */
            });
            logger.info("PrototypeBOM get master BOM stop");
        } 
        catch (err) 
        {
            logger.error("Can't get master BOM from WC" + err);
        }

        //step 2: Create string of all Ref des Which Are CCritical
        //get All C-Crititcal Components From ECAD BOM
        params = {
            fieldName: "componentClass" /* STRING */ ,
            isCaseSensitive: true /* BOOLEAN */ ,
            t: masterBom /* INFOTABLE */ ,
            value: "CCRTL" /* STRING */
        };
        logger.info("PrototypeBOM filter CCTRLcomponetnts " + masterBom.rows.length);
        ecadcCRTLTable = Resources["InfoTableFunctions"].EQFilter(params);

        // OPT-1: Build hash set for O(1) CCRTL refDes lookup instead of string indexOf
        var cCRTLSet = {};
        cCRTLString = "";
        logger.info("PrototypeBOM start loop on cctrTbl");
        for (i = 0; i < ecadcCRTLTable.rows.length; i++) {
            ecadcCRTLRow = ecadcCRTLTable.rows[i];
            cCRTLString = ecadcCRTLRow.refDes + "," + cCRTLString;
            cCRTLSet[ecadcCRTLRow.refDes] = true;
        }

        // If C-CRTL == YES, then add to a  cCriticalTable.  cCriticalTable is sent to PPTA later.
        cCriticalTable = Resources["InfoTableFunctions"].CreateInfoTableFromDataShape({
            infoTableName: "InfoTable",
            dataShapeName: "WHR.IsCCriticalOutputDS_AtosSyntel"
        });

        params = {
            infoTableName: "InfoTable",
            dataShapeName: "WHR.VariantBomCompareReportDS_AtosSyntel"
        };

        consolidatedBom = Resources["InfoTableFunctions"].CreateInfoTableFromDataShape(params);

        params = {
            infoTableName: "InfoTable",
            dataShapeName: "WHR.VariantBomCompareReportDS_AtosSyntel"
        };

        tempTable = Resources["InfoTableFunctions"].CreateInfoTableFromDataShape(params);

        params = {
            infoTableName: "InfoTable",
            dataShapeName: "WHR.PartVariantDataSelectionDS_AtosSyntel"
        };

        selectedVariantTable = Resources["InfoTableFunctions"].CreateInfoTableFromDataShape(params);

        for (i = 0; i < allSelectedVariantsTable.rows.length; i++) {
            newEntry = new Object();
            variantRow = allSelectedVariantsTable.rows[i];
            newEntry.oid = variantRow.oid; // STRING
            newEntry.number = variantRow.number;

            selectedVariantTable.AddRow(newEntry);
        }

        logger.info("PrototypeBOM generate dynamicInfotable start");
        overpopulatedBOM = Things["WHR.PrototypeBomThing_AtosSyntel"].getDynamicInfotable({
            allVariantTable: selectedVariantTable /* INFOTABLE */
        });
        logger.info("PrototypeBOM generate dynamicInfotable stop and addallinfotablesin one table start");
        allBomsInSingleTable = Things["WHR.PrototypeBomThing_AtosSyntel"].addAllInfotablesInOneTable({
            selectedVariantInfotable: selectedVariantTable /* INFOTABLE */
        });
        logger.info("PrototypeBOM generate dynamicInfotable stop and addallinfotablesin one table stop");
        allConsolidatedBomsTableLength = allBomsInSingleTable.rows.length;

        for (i = 0; i < allConsolidatedBomsTableLength; i++) {
            allBomsRow = allBomsInSingleTable.rows[i];
            variantBomMain = allBomsRow.variantTable;

            for (j = 0; j < variantBomMain.rows.length; j++) {
                newEntry = new Object();

                variantRow = variantBomMain.rows[j];

                newEntry.whirlpoolP_N = variantRow.whirlpoolP_N; // STRING
                newEntry.refDes = variantRow.refDes; // STRING
                newEntry.description = variantRow.description; // STRING
                //Overwrite CCritical component Class
                newEntry.componentClass = variantRow.componentClass; // STRING 

                // OPT-1: Use hash set instead of cCRTLString.indexOf()
                if (cCRTLSet[variantRow.refDes] === true) {
                    newEntry.componentClass = "CCRTL";
                }
                //newEntry.PPL = variantRow.PPL; // STRING
                newEntry.oid = variantRow.oid;
                consolidatedBom.AddRow(newEntry);
            }
        }


        // OPT-2: Replace Distinct + Query loop with single-pass hash map
        // Original called Distinct() then Query() for each distinct row — O(distinct × total).
        // Now a single pass keeps the first row per unique (whirlpoolP_N, componentClass) key.
        logger.info("PrototypeBOM consolidatedBOM " + consolidatedBom.getRows());
        var distinctMap = {};
        for (i = 0; i < consolidatedBom.rows.length; i++) {
            row = consolidatedBom.rows[i];
            var dKey = row.whirlpoolP_N + "|" + row.componentClass;
            if (!distinctMap[dKey]) {
                distinctMap[dKey] = true;
                tempTable.AddRow(row);
            }
        }

        // OPT-3: Pre-apply CCRTL overwrite to all variant BOM rows (one-time pass)
        // and pre-build per-variant lookup maps for O(1) access in main loop.
        // Original re-applied CCRTL overwrite and called Query() inside O(parts × variants) loop.
        var variantBomMaps = [];
        for (i = 0; i < allConsolidatedBomsTableLength; i++) {
            allBomsRow = allBomsInSingleTable.rows[i];
            var vBom = allBomsRow.variantTable;
            var vMap = {};

            for (var n = 0; n < vBom.rows.length; n++) {
                variantRow = vBom.rows[n];
                // Apply CCRTL overwrite once (original did this redundantly every outer iteration)
                if (cCRTLSet[variantRow.refDes] === true) {
                    variantRow.componentClass = "CCRTL";
                }
                var vKey = variantRow.whirlpoolP_N + "|" + variantRow.componentClass;
                if (!vMap[vKey]) {
                    vMap[vKey] = [];
                }
                vMap[vKey].push(variantRow);
            }

            variantBomMaps.push({
                variantId: allBomsRow.variantId,
                map: vMap
            });
        }

        // OPT-4: Pre-build masterBom lookup map for totalPopQty — O(1) instead of EQFilter per part
        var masterBomQtyMap = {};
        for (i = 0; i < masterBom.rows.length; i++) {
            row = masterBom.rows[i];
            if (!masterBomQtyMap[row.whirlpoolP_N]) {
                masterBomQtyMap[row.whirlpoolP_N] = row.quantity;
            }
        }

        for (x = 0; x < tempTable.rows.length; x++) {
            consolidatedBomRow = tempTable.rows[x];

            newEntry = new Object();
            var lookupKey = consolidatedBomRow.whirlpoolP_N + "|" + consolidatedBomRow.componentClass;

            for (i = 0; i < allConsolidatedBomsTableLength; i++) {
                // OPT-3: Use pre-built variant BOM maps instead of Query()
                var matchingRows = variantBomMaps[i].map[lookupKey];

                if (!matchingRows || matchingRows.length <= 0) {
                    variantRefField = variantBomMaps[i].variantId + " RefDes";
                    newEntry[variantRefField] = "";
                    variantQtyField = variantBomMaps[i].variantId + " Qty";
                    newEntry[variantQtyField] = "";
                } else {
                    refDesCSV = " ";

                    // Sort matching rows by refDes ascending to match original InfoTable Sort behavior
                    matchingRows.sort(function(a, b) {
                        var aVal = a.refDes || "";
                        var bVal = b.refDes || "";
                        return aVal < bVal ? -1 : aVal > bVal ? 1 : 0;
                    });

                    for (j = 0; j < matchingRows.length; j++) {
                        samePartExhistTableRow = matchingRows[j];

                        if (samePartExhistTableRow.refDes !== undefined) {
                            //>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<
                            for (z = 1; z < samePartExhistTableRow.refDes.length; z++) {
                                var fnd = samePartExhistTableRow.refDes.indexOf(",", z);

                                if (fnd !== -1) {
                                    var str = samePartExhistTableRow.refDes.substring(z, fnd);

                                    if (refDesCSV.indexOf(str) == -1) {
                                        refDesCSV += str;

                                    }
                                    z += fnd - 1;
                                } else if (z < 2) {
                                    refDesCSV = refDesCSV.concat(samePartExhistTableRow.refDes);
                                }
                            }
                            refDesCSV = refDesCSV.concat(", ");
                        }
                    }

                    // There is ", " at the end of CSV.  Remove it.
                    if (refDesCSV.length > 2) {
                        refDesCSV = refDesCSV.substring(0, refDesCSV.length - 2);
                    }

                    //comma seperated alphanumeric string to array
                    refArray = refDesCSV.split(",");

                    //Total Population quantity on Refdes count
                    if (refArray.length !== 0) {
                        refDesForTPQ = refArray.length;
                    }

                    // call a sort function 
                    refArray.sort(sortAlphaNum);
                    refDesCSV = refArray.toString();

                    variantRefField = variantBomMaps[i].variantId + " RefDes";
                    newEntry[variantRefField] = refDesCSV;
                    variantQtyField = variantBomMaps[i].variantId + " Qty";
                    //newEntry[variantQtyField] = samePartExhistTable.rows[0].quantity;
                    refDesForTPQ = refDesForTPQ.toString();
                    newEntry[variantQtyField] = refDesForTPQ.split(".")[0];

                    // insert the ref des for column sorting
                    var sortOrderFlag = false;
                    if (refArray.length > 0 && sortOrderFlag === false) {
                        sortOrderFlag = true;
                        refDes = refDesCSV.split(",")[0];
                        part1 = refDes.match(/[a-zA-Z]+/);
                        part2 = refDes.match(/[0-9]+/);
                        newEntry.sortOrderRefDes = part1 + ("0000" + part2).substr(-4);
                    }
                }
            }

            newEntry.whirlpoolP_N = consolidatedBomRow.whirlpoolP_N; // STRING
            newEntry.description = consolidatedBomRow.description; // STRING
            newEntry.componentClass = consolidatedBomRow.componentClass; // STRING
            newEntry.oid = consolidatedBomRow.oid;
            //logger.info(" consolidatedBomRow " + consolidatedBomRow.whirlpoolP_N + "," + consolidatedBomRow.componentClass + "," + consolidatedBomRow.oid);

            // OPT-4: Use pre-built masterBom map instead of EQFilter per part
            if (masterBomQtyMap[consolidatedBomRow.whirlpoolP_N] !== undefined) {
                newEntry.totalPopQty = masterBomQtyMap[consolidatedBomRow.whirlpoolP_N];
            }

            // add cCritical parts in a seprate table 
            if (consolidatedBomRow.componentClass === "CCRTL") {
                newEntry3 = new Object();
                newEntry3.componentsPartNumber = consolidatedBomRow.whirlpoolP_N; // STRING
                newEntry3.ecadPartNumber = EcadNumber; // STRING
                newEntry3.refDes = consolidatedBomRow.refDes; // STRING
                cCriticalTable.AddRow(newEntry3);
            }
            // All Windchill data afinal table 
            overpopulatedBOM.AddRow(newEntry);
        }

        //get ppta record and store in table 
        sort = new Object();
        sort.name = "refDes";
        sort.ascending = true;
        cCriticalTable.Sort(sort);

        sort = new Object();
        sort.name = "sortOrderRefDes";
        sort.ascending = true;
        overpopulatedBOM.Sort(sort);

        try {
            // Get data from PPTA
            pptaResponseJSON = Things["WHR.BomAndRiskToolsODataConnectorThing_AtosSyntel"].getManufacturerDetails({
                InputString: cCriticalTable.ToJSON().rows /* STRING */
            });

            logger.info("Response from ppta for prototype : " + pptaResponseJSON.statusCode);
            logger.info("Input string for prototype : " + cCriticalTable.ecadPartNumber + "," + cCriticalTable.componentsPartNumber + "," + cCriticalTable.refDes);
            logger.info("ppta full input string: " + cCriticalTable.ToJSON().rows);
            logger.info("ppta response " +pptaResponseJSON.rows);

            pptaResponseTable = Resources["InfoTableFunctions"].CreateInfoTableFromDataShape({
                infoTableName: "InfoTable",
                dataShapeName: "WHR.PPTAReportDS_AtosSyntel"
            });

            if (pptaResponseJSON.pptaNavigateResponseDtoList !== undefined) {
                for (i = 0; i < pptaResponseJSON.pptaNavigateResponseDtoList.length; i++) {
                    row = pptaResponseJSON.pptaNavigateResponseDtoList[i];

                    // WHR.PPTAReportDS_AtosSyntel entry object
                    newEntry2 = new Object();
                    newEntry2.Status = row.status; // STRING
                    newEntry2.ManufacturerName = row.manufacturerName; // STRING
                    newEntry2.RefDes = row.refDes; // STRING
                    newEntry2.ComponentPartNumber = row.componentPartNumber; // STRING
                    newEntry2.Mpn = row.mpn; // STRING
                    newEntry2.EcadPartNumber = row.ecadPartNumber; // STRING
                    pptaResponseTable.AddRow(newEntry2);
                }
            }
        } catch (err) {
            logger.error("Cant get data from PPTA " + err);
        }

        // OPT-5: Pre-index pptaResponseTable by ComponentPartNumber for O(1) lookup
        // Replaces per-row EQFilter calls in the AXL loop below
        var pptaByPartNumber = {};
        if (pptaResponseTable !== undefined) {
            for (i = 0; i < pptaResponseTable.rows.length; i++) {
                row = pptaResponseTable.rows[i];
                if (!pptaByPartNumber[row.ComponentPartNumber]) {
                    pptaByPartNumber[row.ComponentPartNumber] = [];
                }
                pptaByPartNumber[row.ComponentPartNumber].push(row);
            }
        }

        // OPT-6 & OPT-7: Build axlEntriesMap directly in a single pass without intermediate InfoTable overhead
        var axlCache = {};
        var axlEntriesMap = {};

        for (x = 0; x < overpopulatedBOM.rows.length; x++) {
            row1 = overpopulatedBOM.rows[x];
            var mapKey = row1.whirlpoolP_N + "|" + row1.componentClass;
            var pptaMatches = (row1.componentClass === "CCRTL") ? pptaByPartNumber[row1.whirlpoolP_N] : null;

            if (pptaMatches && pptaMatches.length > 0) {
                if (supplierCount < pptaMatches.length) {
                    supplierCount = pptaMatches.length;
                }

                var pptaList = [];
                for (k = 0; k < pptaMatches.length; k++) {
                    row3 = pptaMatches[k];
                    pptaList.push({
                        MPN: row3.Mpn,
                        manufaturerName: row3.ManufacturerName,
                        partNumber: row1.whirlpoolP_N,
                        componentClass: row1.componentClass
                    });
                }
                axlEntriesMap[mapKey] = pptaList;
            } else {
                var axlCacheKey = row1.oid + "|" + row1.whirlpoolP_N;
                if (axlCache[axlCacheKey] === undefined) {
                    axlCache[axlCacheKey] = Things["WHR.PrototypeBomThing_AtosSyntel"].getAxlEntries({
                        partOid: row1.oid /* STRING */ ,
                        partNumber: row1.whirlpoolP_N
                    });
                }
                axlTable = axlCache[axlCacheKey];

                if (axlTable !== undefined && axlTable.rows.length > 0) {
                    if (supplierCount < axlTable.rows.length) {
                        supplierCount = axlTable.rows.length;
                    }

                    var axlList = [];
                    for (z = 0; z < axlTable.rows.length; z++) {
                        row = axlTable.rows[z];
                        axlList.push({
                            MPN: row.MPN,
                            manufaturerName: row.manufaturerName,
                            WCPartOid: row.WCPartOid,
                            partNumber: row1.whirlpoolP_N,
                            componentClass: row1.componentClass
                        });
                    }
                    axlEntriesMap[mapKey] = axlList;
                }
            }
        }

        finalTable = Things["WHR.PrototypeBomThing_AtosSyntel"].getDynamicInfotableFinalTable({
            supplierCount: supplierCount /* INTEGER */ ,
            allVariantTable: allSelectedVariantsTable /* INFOTABLE */
        });

        if ((overpopulatedBOM.dataShape === null) || (overpopulatedBOM.dataShape === undefined)) {
            iLF = overpopulatedBOM.ToJSON().dataShape.fieldDefinitions;
        } else {
            iLF = overpopulatedBOM.dataShape.fields;
        }

        for (x = 0; x < overpopulatedBOM.rows.length; x++) {
            sourceRow = overpopulatedBOM.rows[x];
            newEntry = new Object();

            for (keyC in iLF) {
                newEntry[keyC] = sourceRow[keyC];
            }

            // OPT-7: Use pre-indexed AXL entries instead of Query() per row
            var axlLookupKey = sourceRow.whirlpoolP_N + "|" + sourceRow.componentClass;
            var axlMatches = axlEntriesMap[axlLookupKey] || [];

            countAxl = 1;
            for (i = 0; i < axlMatches.length; i++) {
                row1 = axlMatches[i];

                supplierField = "Supplier " + countAxl;
                newEntry[supplierField] = row1.manufaturerName;
                //supplierPartField = "Supplier Part Number " + countAxl;
                supplierPartField = "Supplier " + countAxl + " Part Number";
                newEntry[supplierPartField] = row1.MPN;
                countAxl++;
            }

            finalTable.AddRow(newEntry);
        }
    }

    if (pptaResponseJSON.statusCode !== "200") {
        newEntry = new Object();
        newEntry.description = "Warning: No response from PPTA"; // STRING
        finalTable.AddRow(newEntry);
    } else if (pptaResponseTable.rows.length === 0) {
        newEntry = new Object();
        newEntry.description = "Note: No CCRTL data found in PPTA"; // STRING
        finalTable.AddRow(newEntry);
    }
    result = finalTable;
} 
catch (err) 
{
    let errMsg = "Thing [{}] Service [{}] error at line [{}] : {}";
    logger.error(errMsg, me.name, err.fileName, err.lineNumber, err);
}
