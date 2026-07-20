try {
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

	function sortAlphaNum(a, b) {
		aA = a.replace(reA, "");
		bA = b.replace(reA, "");

		if (aA === bA) {
			aN = parseInt(a.replace(reN, ""), 10);
			bN = parseInt(b.replace(reN, ""), 10);

			return aN === bN ? 0 : aN > bN ? 1 : -1;
		} else {
			return aA > bA ? 1 : -1;
		}
	}

	params = {
		infoTableName: "finalComparisonTable" /* STRING */
	};

	finalTable = Resources["InfoTableFunctions"].CreateInfoTable(params);


	if (masterOid !== undefined && masterOid !== "") {
		try {
			//get Master Bom
            logger.info("PrototypeBOM get master BOM start");
			masterBom = Things["WHR.PrototypeBomThing_AtosSyntel"].getBomInfotableForPrototypeBom({
				WTPartOid: masterOid /* STRING */
			});
             logger.info("PrototypeBOM get master BOM stop");
		} catch (err) {
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

		cCRTLString = "";
logger.info("PrototypeBOM start loop on cctrTbl");
		for (i = 0; i < ecadcCRTLTable.rows.length; i++) {
			ecadcCRTLRow = ecadcCRTLTable.rows[i];
			cCRTLString = ecadcCRTLRow.refDes + "," + cCRTLString;
			//zb
			//        logger.info("ecadinfo " + ecadcCRTLTable.rows[i]);
		}
		//zb
		//    logger.info("Crtitical info :" + cCRTLString );

		//Result thill here -> U8,U6,U7,U9,U2,U3,U15,U13

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

				if (cCRTLString.indexOf(variantRow.refDes + ",") !== -1) {
					newEntry.componentClass = "CCRTL";
				}
				//newEntry.PPL = variantRow.PPL; // STRING
				newEntry.oid = variantRow.oid;
				consolidatedBom.AddRow(newEntry);
			}
		}


		params = {
			t: consolidatedBom /* INFOTABLE */ ,
			columns: "whirlpoolP_N,componentClass" /* STRING */
		};
logger.info("PrototypeBOM consolidatedBOM " + consolidatedBom.getRows());
		distinctConsolidatedBom = Resources["InfoTableFunctions"].Distinct(params);

		//
		//  testcsvipnut2 = Resources["InfotableFunctions"].Clone({ t1: distinctConsolidatedBom });   
		//>>>>>>>>>>>>>>//

		for (i = 0; i < distinctConsolidatedBom.rows.length; i++) {
			//        params =
			//        {
			//            fieldName: "whirlpoolP_N" /* STRING */,
			//            isCaseSensitive: undefined /* BOOLEAN */,
			//            t: consolidatedBom /* INFOTABLE */,
			//            value: distinctConsolidatedBom.rows[i].whirlpoolP_N /* STRING */
			//        };
			//
			////        tempResult = Resources["InfoTableFunctions"].EQFilter(params);
			//	
			query = {
				"filters": {
					"type": "And",
					"filters": [{
							"type": "EQ",
							"fieldName": "whirlpoolP_N",
							"value": distinctConsolidatedBom.rows[i].whirlpoolP_N
						},
						{
							"type": "EQ",
							"fieldName": "componentClass",
							"value": distinctConsolidatedBom.rows[i].componentClass
						}
					]
				}
			};

			params = {
				t: consolidatedBom /* INFOTABLE */ ,
				query: query /* QUERY */
			};

			tempResult = Resources["InfoTableFunctions"].Query(params);
			//>>>>>
			//   if( tempResult !== undefined && tempResult !== "" ) {   
			//   testcsvipnut = Resources["InfotableFunctions"].Clone({ t1: tempResult }); 
			//   }
			//>>>>>      
			if (tempResult.rows.length > 0) {
				tempTable.AddRow(tempResult.rows[0]);
			}
		}



		for (x = 0; x < tempTable.rows.length; x++) {
			consolidatedBomRow = tempTable.rows[x];

			newEntry = new Object();

			for (i = 0; i < allConsolidatedBomsTableLength; i++) {
				allBomsRow = allBomsInSingleTable.rows[i];
				variantBom = allBomsRow.variantTable;

				query = {
					"filters": {
						"type": "And",
						"filters": [{
								"type": "EQ",
								"fieldName": "whirlpoolP_N",
								"value": consolidatedBomRow.whirlpoolP_N
							},
							{
								"type": "EQ",
								"fieldName": "componentClass",
								"value": consolidatedBomRow.componentClass
							}
						]
					}
				};

				params = {
					t: variantBom /* INFOTABLE */ ,
					query: query /* QUERY */
				};

				//>>>>>
				//   testcsvipnut = Resources["InfotableFunctions"].Clone({ t1: tempTable }); 
				//>>>>>

				samePartExhistTable = Resources["InfoTableFunctions"].Query(params);
				//            params =
				//            {
				//                fieldName: "whirlpoolP_N" /* STRING */,
				//                isCaseSensitive: undefined /* BOOLEAN */,
				//                t: variantBom /* INFOTABLE */,
				//                value: consolidatedBomRow.whirlpoolP_N /* STRING */
				//            };
				//
				//            samePartExhistTable = Resources["InfoTableFunctions"].EQFilter(params);

				if (samePartExhistTable.rows.length <= 0) {
					variantRefField = allBomsRow.variantId + " RefDes";
					newEntry[variantRefField] = "";
					variantQtyField = allBomsRow.variantId + " Qty";
					newEntry[variantQtyField] = "";
				} else {
					refDesCSV = " ";

					// Sort RefDes table that CSV is also sorted
					sort = new Object();
					sort.name = "refDes";
					sort.ascending = true;
					samePartExhistTable.Sort(sort);

					for (j = 0; j < samePartExhistTable.rows.length; j++) {
						samePartExhistTableRow = samePartExhistTable.rows[j];

						if (samePartExhistTableRow.refDes !== undefined) {
							//>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<
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


							//    return occurrencesCSV;
							//>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<< 
							//>>>                       refDesCSV = refDesCSV.concat(samePartExhistTableRow.refDes);
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

					variantRefField = allBomsRow.variantId + " RefDes";
					newEntry[variantRefField] = refDesCSV;
					variantQtyField = allBomsRow.variantId + " Qty";
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
			params = {
				fieldName: "whirlpoolP_N" /* STRING */ ,
				isCaseSensitive: undefined /* BOOLEAN */ ,
				t: masterBom /* INFOTABLE */ ,
				value: consolidatedBomRow.whirlpoolP_N /* STRING */
			};

			totalPopulationQtyTable = Resources["InfoTableFunctions"].EQFilter(params);


			if (totalPopulationQtyTable.rows.length > 0) {
				newEntry.totalPopQty = totalPopulationQtyTable.rows[0].quantity;
			}

			// newEntry.totalPopQty = refDesForTPQ;


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

		// add all axl entries in single table
		params = {
			infoTableName: "InfoTable",
			dataShapeName: "WHR.ManufaturerDetailsFromAxlDS_AtosSyntel"
		};

		allAxlEntriesTable = Resources["InfoTableFunctions"].CreateInfoTableFromDataShape(params);

		for (x = 0; x < overpopulatedBOM.rows.length; x++) {
			row1 = overpopulatedBOM.rows[x];

			//check if part is ccritical
			if (row1.componentClass === "CCRTL") {
				params = {
					fieldName: "ComponentPartNumber" /* STRING */ ,
					isCaseSensitive: undefined /* BOOLEAN */ ,
					t: pptaResponseTable /* INFOTABLE */ ,
					value: row1.whirlpoolP_N /* STRING */
				};

				filteredPptaResponseTable = Resources["InfoTableFunctions"].EQFilter(params);

				if (filteredPptaResponseTable.rows.length > 0) {
					if (supplierCount < filteredPptaResponseTable.rows.length) {
						supplierCount = filteredPptaResponseTable.rows.length;
					}

					for (k = 0; k < filteredPptaResponseTable.rows.length; k++) {
						newEntry1 = new Object();
						row3 = filteredPptaResponseTable.rows[k];

						newEntry1.MPN = row3.Mpn; // STRING
						newEntry1.manufaturerName = row3.ManufacturerName; // STRING
						newEntry1.partNumber = row1.whirlpoolP_N;
						newEntry1.componentClass = row1.componentClass;
						allAxlEntriesTable.AddRow(newEntry1);
					}
				} else {
					axlTable = Things["WHR.PrototypeBomThing_AtosSyntel"].getAxlEntries({
						partOid: row1.oid /* STRING */ ,
						partNumber: row1.whirlpoolP_N
					});

					if (axlTable !== undefined) {
						if (axlTable.rows.length > 0) {
							if (supplierCount < axlTable.rows.length) {
								supplierCount = axlTable.rows.length;
							}

							// Add Supplier and axl entries in single table 
							for (z = 0; z < axlTable.rows.length; z++) {
								row = axlTable.rows[z];
								newEntry4 = new Object();
								newEntry4.MPN = row.MPN; // STRING
								newEntry4.manufaturerName = row.manufaturerName; // STRING
								newEntry4.WCPartOid = row.WCPartOid; // STRING
								newEntry4.partNumber = row1.whirlpoolP_N;
								newEntry4.componentClass = row1.componentClass;
								allAxlEntriesTable.AddRow(newEntry4);
							}
						}
					}
				}
			} else {
				axlTable = Things["WHR.PrototypeBomThing_AtosSyntel"].getAxlEntries({
					partOid: row1.oid /* STRING */ ,
					partNumber: row1.whirlpoolP_N
				});

				if (axlTable !== undefined) {
					if (axlTable.rows.length > 0) {
						if (supplierCount < axlTable.rows.length) {
							supplierCount = axlTable.rows.length;
						}

						// Add Supplier and axl entries in single table 
						for (z = 0; z < axlTable.rows.length; z++) {
							row = axlTable.rows[z];
							newEntry4 = new Object();
							newEntry4.MPN = row.MPN; // STRING
							newEntry4.manufaturerName = row.manufaturerName; // STRING
							newEntry4.WCPartOid = row.WCPartOid; // STRING
							newEntry4.partNumber = row1.whirlpoolP_N;
							newEntry4.componentClass = row1.componentClass;
							allAxlEntriesTable.AddRow(newEntry4);
						}
					}
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

			//query Axl entries table to get all supplier for current part. 
			//        params =
			//        {
			//            fieldName: "partNumber" /* STRING */,
			//            isCaseSensitive: undefined /* BOOLEAN */,
			//            t: allAxlEntriesTable /* INFOTABLE */,
			//            value: sourceRow.whirlpoolP_N /* STRING */
			//        };
			//
			//        axlSamePartExhistTable = Resources["InfoTableFunctions"].EQFilter(params);

			query = {
				"filters": {
					"type": "And",
					"filters": [{
							"type": "EQ",
							"fieldName": "partNumber",
							"value": sourceRow.whirlpoolP_N
						},
						{
							"type": "EQ",
							"fieldName": "componentClass",
							"value": sourceRow.componentClass
						}
					]
				}
			};

			params = {
				t: allAxlEntriesTable /* INFOTABLE */ ,
				query: query /* QUERY */
			};

			axlSamePartExhistTable = Resources["InfoTableFunctions"].Query(params);

			countAxl = 1;
			for (i = 0; i < axlSamePartExhistTable.rows.length; i++) {
				row1 = axlSamePartExhistTable.rows[i];

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
} catch (err){
	let errMsg = "Thing [{}] Service [{}] error at line [{}] : {}";
	logger.error(errMsg, me.name, err.fileName, err.lineNumber, err);
}