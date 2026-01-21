sap.ui.define(["sap/ui/core/mvc/ControllerExtension", "xlsx"], 
	
	function (ControllerExtension, XLSX) {
	'use strict';

	return ControllerExtension.extend('re.airline.ext.controller.ListReportExt', {
		// this section allows to extend lifecycle hooks or hooks provided by Fiori elements
		override: {
			/**
             * Called when a controller is instantiated and its View controls (if available) are already created.
             * Can be used to modify the View before it is displayed, to bind event handlers and do other one-time initialization.
             * @memberOf re.airline.ext.controller.ListReportExt
             */
			onInit: function () {
				// you can access the Fiori elements extensionAPI via this.base.getExtensionAPI
				var oModel = this.base.getExtensionAPI().getModel();
			}
		},

		excelSheetsData: [],
        pDialog: null,
		
		openExcelUploadDialog: function () {
			console.log(XLSX.version)			
			var oView = this.getView();
			var sFragId = oView.createId("excel_upload");
			if (!this.pDialog) {
				sap.ui.core.Fragment.load({
				id: sFragId,
				name: "re.airline.ext.fragment.ExcelUpload",
				type: "XML",
				controller: this
				}).then((oDialog) => {
				oView.addDependent(oDialog);

				var oUploadSet = sap.ui.core.Fragment.byId(sFragId, "uploadSet");
				if (oUploadSet) { oUploadSet.removeAllItems(); }

				this.pDialog = oDialog;
				this.pDialog.open();
				}).catch((error) => {
				sap.m.MessageBox.error("Failed to load Excel Upload fragment:\n" + error.message);
				});
			} else {
				var oUploadSet = sap.ui.core.Fragment.byId(sFragId, "uploadSet");
				if (oUploadSet) { oUploadSet.removeAllItems(); }
				this.pDialog.open();
			}
		},

		onUploadSet: function (oEvent) {
			// 1) rows guard
			const rows =
				this.excelRows
				|| this.excelSheetsDataNormalized?.rows
				|| this.excelSheetsData?.Sheet1
				|| [];
			if (!Array.isArray(rows) || rows.length === 0) {
				sap.m.MessageToast.show("Select file to Upload");
				return;
			}

			// 2) resolve FE V4 API + EditFlow
			const extApi = this.base?.getExtensionAPI?.();
			const editFlow = extApi?.getEditFlow?.();

			// 3) prevent double submit
			const btn = oEvent?.getSource?.();
			btn?.setEnabled?.(false);

			// 4) wrap your create loop
			const task = () => new Promise((resolve, reject) => {
				try {
				// pass rows if your callOdata accepts them
				// this.callOdata(resolve, reject, rows);
				this.callOdata(resolve, reject);
				} catch (e) { reject(e); }
			});

			const onDone = () => {
				btn?.setEnabled?.(true);
			};

			// 5) prefer EditFlow.secureExecution (or securedExecution on some stacks)
			const secureExec = editFlow?.secureExecution || editFlow?.securedExecution;

			if (typeof secureExec === "function") {
				secureExec.call(editFlow, task, {
				sActionLabel: btn?.getText?.() || "Upload",
				busy: { set: true, check: true },
				dataloss: { popup: false },
				action: true
				})
				.then(() => {
				sap.m.MessageToast.show("Upload completed.");
				this._cleanupUploadDialog?.();
				})
				.catch(err => sap.m.MessageBox.error("Upload failed:\n" + (err?.message || err)))
				.finally(onDone);
			} else {
				// 6) fallback if EditFlow/secureExecution not available
				this.getView().setBusy(true);
				task()
				.then(() => {
					sap.m.MessageToast.show("Upload completed.");
					this._cleanupUploadDialog?.();
				})
				.catch(err => sap.m.MessageBox.error("Upload failed:\n" + (err?.message || err)))
				.finally(() => {
					this.getView().setBusy(false);
					onDone();
				});
			}
		},

			// Optional helper to clear dialog + file + in-memory rows
			_cleanupUploadDialog: function () {
			this.excelRows = [];
			if (this.excelSheetsDataNormalized) this.excelSheetsDataNormalized.rows = [];

			const oUploadSet = this.byId("uploadSet") || sap.ui.getCore().byId("excel_upload--uploadSet");
			if (oUploadSet) oUploadSet.removeAllItems();

			if (this.pDialog?.close) this.pDialog.close();
		},

        onTemplateDownload: function (oEvent) {
            console.log("Template Download Button Clicked!!!")

            // get the odata model binded to this application
            var oModel = this.getView().getModel();

			// Ensure oModel is your v4 ODataModel, e.g. this.getView().getModel()
			const oMetaModel = oModel.getMetaModel();

			// 1) Get the qualified type name behind the entity set
			oMetaModel.requestObject("/AeroRegistry/$Type").then(function (sQualifiedType) {
				// 2) Load the entity type metadata object
				return oMetaModel.requestObject("/" + sQualifiedType + "/");
			}).then(function (oEntityType) {

				const propertyList = [
					'Airlinecode',
					'Rpcid',
					'Iatacode',
					'Aircrafttype',
					'Mtowquantity',
					'Mtowunit'
				];

				// 5) Map to column headers (prefer @Common.Label, fallback to property name)
				const aHeaders = propertyList.map((sProp) => {
				const oProp = oEntityType[sProp];
				if (oProp && oProp.$kind === "Property") {
					// V4 annotations: @com.sap.vocabularies.Common.v1.Label
					return oProp["@com.sap.vocabularies.Common.v1.Label"] || sProp;
				}
				// Not found in metadata → still add the property name so the template remains usable
				return sProp;
				});

				// 6) Build header-only sheet (AOA ensures only headers without dummy rows)
				const ws = XLSX.utils.aoa_to_sheet([aHeaders]);

				// Optional: adjust column widths (auto width-ish)
				ws["!cols"] = aHeaders.map(h => ({ wch: Math.max(14, String(h).length + 2) }));

				// 7) Create workbook and append the sheet
				const wb = XLSX.utils.book_new();
				XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

				// 8) Download
				XLSX.writeFile(wb, "AeroRegistry_Template.xlsx");

				MessageToast.show("Template file downloading…");

			}).catch(function (err) {
				console.error("Failed to read V4 metadata:", err);
			});

        },

        onCloseDialog: function (oEvent) {
            this.pDialog.close();
        },
        onBeforeUploadStart: function (oEvent) {
            console.log("File Before Upload Event Fired!!!")
            /* TODO: check for file upload count */
        },
    // // === MTOW VALIDATION FUNCTION ===
    // 	isInvalidMtow: function (value) {
	// 		if (value === "" || value == null) return true;   // cannot be empty

	// 		const n = Number(value);
	// 		if (isNaN(n)) return true;                        // must be numeric
	// 		if (n <= 0) return true;                          // must be > 0

	// 		// Recommended realistic aviation range
	// 		if (n < 100 || n > 1000000) return true;

	// 		return false;
    // 	},
		onUploadSetComplete: async function (oEvent) {
			// === CONFIG ===
			const REQUIRED_PROPS = [
					'Airlinecode',
					'Rpcid',
					'Iatacode',
					'Aircrafttype',
					'Mtowquantity',
					'Mtowunit'
			];
			const ENTITY_SET = "AeroRegistry"; //
			// =============
			// === MTOW VALIDATION FUNCTION ===
			function isInvalidMtow(value) {
				if (value === "" || value == null) return true;   // cannot be empty

				const n = Number(value);
				if (isNaN(n)) return true;                        // must be numeric
				if (n <= 0) return true;                          // must be > 0
				return false;
			}

			try {
				// const oUploadSet = sap.ui.core.Fragment.byId("excel_upload", "uploadSet");
				// const oUploadSet = this._getUploadSet();
				const oUploadSet = oEvent.getSource();
				const aItems = oUploadSet ? oUploadSet.getItems() : [];
				if (!aItems.length) {
				sap.m.MessageToast.show("No file selected.");
				return;
				}

				const oItem = aItems[0];
				const oFile = await oItem.getFileObject();
				if (!oFile) {
				sap.m.MessageToast.show("Could not access the selected file.");
				return;
				}

				this.getView().setBusy(true);

				// --- Read workbook as ArrayBuffer
				const reader = new FileReader();
				reader.onload = async (e) => {
				try {
					const wb = XLSX.read(e.target.result, { type: "array" });

					// Prefer Sheet1, else first
					const sheetName = wb.Sheets["Sheet1"] ? "Sheet1" : (wb.SheetNames[0] || null);
					if (!sheetName) throw new Error("No sheets found in the workbook.");

					// Read as AOA to take full control of header mapping
					const aoa = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], {
					header: 1, // first row = headers
					blankrows: false,
					defval: "" // keep empty cells as empty string
					});

					if (!aoa.length) throw new Error("The sheet is empty.");
					const headerRow = (aoa[0] || []).map(h => String(h || "").trim());
					const dataRows = aoa.slice(1);

					// --- Build label<->property maps from V4 metadata
					const oModel = this.getView().getModel();
					const oMetaModel = oModel.getMetaModel();

					const sQualifiedType = await oMetaModel.requestObject("/" + ENTITY_SET + "/$Type");
					const oEntityType = await oMetaModel.requestObject("/" + sQualifiedType + "/");

					// Build maps:
					// propName -> label, and label (case-insensitive) -> propName
					const labelByProp = {};
					const propByLabelLC = {};
					Object.keys(oEntityType).forEach((k) => {
					const o = oEntityType[k];
					if (o && o.$kind === "Property") {
						const label = o["@com.sap.vocabularies.Common.v1.Label"] || k;
						labelByProp[k] = label;
						propByLabelLC[label.toLowerCase()] = k;
					}
					});

					// Accept headers as either propName or label
					const colIndexByProp = {}; // { PROP: columnIndex }
					headerRow.forEach((hdr, idx) => {
					const lc = hdr.toLowerCase();
					// 1) exact property match
					if (labelByProp[hdr]) {
						colIndexByProp[hdr] = idx;
						return;
					}
					// 2) label match (case-insensitive)
					if (propByLabelLC[lc]) {
						colIndexByProp[propByLabelLC[lc]] = idx;
						return;
					}
					// 3) no-op for unknown columns (ignored)
					});

					// Validate: all REQUIRED_PROPS must be present
					const missing = REQUIRED_PROPS.filter((p) => !(p in colIndexByProp));
					if (missing.length) {
					const humanMissing = missing
						.map((p) => labelByProp[p] || p)
						.join(", ");
					throw new Error("Missing required columns: " + humanMissing);
					}

					// Normalize data rows to fixed shape (REQUIRED_PROPS order)
					const normalizedRows = dataRows.map((row) => {
					const obj = {};
					REQUIRED_PROPS.forEach((p) => {
						const colIdx = colIndexByProp[p];
						obj[p] = colIdx != null ? row[colIdx] : ""; // should exist due to validation
					});
					return obj;
					});
					
					// === VALIDATE MTOWQUANTITY HERE ===
					const invalidMtow = normalizedRows
						.map((r, i) => ({ row: i + 1, value: r.Mtowquantity }))
						.filter(({ value }) => isInvalidMtow(value));

					if (invalidMtow.length > 0) {
						oUploadSet.removeAllItems();
						throw new Error(
							"Invalid Mtowquantity detected:\n" +
							invalidMtow
								.map(e => `Row ${e.row} → "${e.value}"`)
								.join("\n")
						);
					}					
					// === CHECK DUPLICATE RPCID (ignoring blank Rpcid) ===
					const rpcidMap = {}; // key: normalized Rpcid, value: array of row numbers

					normalizedRows.forEach((r, idx) => {
						const raw = r.Rpcid;
						const key = (raw == null ? "" : String(raw)).trim();

						if (!key) {
							// ignore blank Rpcid for duplication check
							return;
						}

						if (!rpcidMap[key]) {
							rpcidMap[key] = [];
						}
						rpcidMap[key].push(idx + 1); // 1-based row number
					});

					const duplicateRpcid = Object.entries(rpcidMap)
						.filter(([_, rows]) => rows.length > 1);

					if (duplicateRpcid.length > 0) {
						oUploadSet.removeAllItems();
						const msg =
							"Duplicate Rpcid detected:\n" +
							duplicateRpcid
								.map(([value, rows]) =>
									`Rpcid "${value}" appears in rows: ${rows.join(", ")}`)
								.join("\n");
						throw new Error(msg);
					}

					// === CHECK RPCID ALREADY EXISTING IN SAP ===
					// Collect distinct, non-empty Rpcid values from Excel
					const aDistinctRpcid = Array.from(
						new Set(
							normalizedRows
								.map(r => (r.Rpcid == null ? "" : String(r.Rpcid)).trim())
								.filter(v => v) // non-empty only
						)
					);

					if (aDistinctRpcid.length > 0) {
						// Build OR filter: Rpcid eq 'A' or Rpcid eq 'B' ...
						const aFilters = aDistinctRpcid.map(v =>
							new sap.ui.model.Filter("Rpcid",
								sap.ui.model.FilterOperator.EQ,
								v)
						);

						const oListBinding = oModel.bindList("/" + ENTITY_SET, undefined, undefined, aFilters);
						const aCtx = await oListBinding.requestContexts(0, Infinity);

						const existingRpcidSet = new Set(
							aCtx
								.map(ctx => ctx.getObject && ctx.getObject().Rpcid)
								.filter(v => v != null)
								.map(v => String(v).trim())
						);

						if (existingRpcidSet.size > 0) {
							// Map which Excel rows have Rpcid that already exist in SAP
							const existingRows = [];

							normalizedRows.forEach((r, idx) => {
								const key = (r.Rpcid == null ? "" : String(r.Rpcid)).trim();
								if (key && existingRpcidSet.has(key)) {
									existingRows.push({
										row: idx + 1,
										value: key
									});
								}
							});

							if (existingRows.length > 0) {
								oUploadSet.removeAllItems();
								const msg =
									"Rpcid already exists in SAP:\n" +
									existingRows
										.map(e => `Row ${e.row} → Rpcid "${e.value}" already exists`)
										.join("\n");
								throw new Error(msg);
							}
						}
					}

					// Store both raw and normalized, if you need
					this.excelSheetsData = this.excelSheetsData || {};
					this.excelSheetsData[sheetName] = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: "", raw: false });

					this.excelSheetsDataNormalized = { rows: normalizedRows, requiredOrder: REQUIRED_PROPS.slice() };

					// Debug
					/* eslint-disable no-console */
					console.log("Headers (raw):", headerRow);
					console.log("Column mapping (prop -> index):", colIndexByProp);
					console.log("Normalized rows:", normalizedRows);

					sap.m.MessageToast.show("Upload successful. Columns validated.");
				} catch (parseErr) {
					sap.m.MessageBox.error("Failed to process Excel:\n" + parseErr.message);
					console.error(parseErr);
				} finally {
					this.getView().setBusy(false);
				}
				};

				reader.onerror = (err) => {
				this.getView().setBusy(false);
				sap.m.MessageBox.error("File read failed.");
				console.error("FileReader error:", err);
				};

				reader.readAsArrayBuffer(oFile);

			} catch (err) {
				this.getView().setBusy(false);
				sap.m.MessageBox.error("Upload handling failed:\n" + err.message);
				console.error(err);
			}
		},
        onItemRemoved:function (oEvent) {
			this.excelSheetsData = [];   
        },

		// Helper method to call OData and create AeroRegistry entries
		callOdata: async function (fnResolve, fnReject) {
			try {
				const oModel = this.getView().getModel(); // sap.ui.model.odata.v4.ODataModel
				const ENTITY_SET = "AeroRegistry";

				const m1 = oModel.getMetaModel();
				const sQType = await m1.requestObject(`/${ENTITY_SET}/$Type`);
				const oEntityType = m1.requestObject(`/${sQType}/`);   // entity type object with properties				

				// Pick rows from your single source of truth
				const rows =
				this.excelRows
				|| this.excelSheetsDataNormalized?.rows
				|| this.excelSheetsData?.Sheet1
				|| [];

				if (!Array.isArray(rows) || rows.length === 0) {
				sap.m.MessageToast.show("No rows to upload.");
				return;
				}

		

				// Use $auto so each create is sent automatically
				const oListBinding = oModel.bindList("/AeroRegistry", undefined, undefined, undefined, {
				$$groupId: "$auto",
				$$updateGroupId: "$auto"
				});

				const mm = sap.ui.getCore().getMessageManager();

				for (let i = 0; i < rows.length; i++) {
					const r = rows[i];
					var uuid = crypto.randomUUID();
					const d = new Date();
					const tzOffset = d.getTimezoneOffset() * 60000;
					const now = new Date().toISOString();
					// Only include properties that exist in your metadata
					const payload = {
					// Uuid: uuid,
					Airlinecode:  r.Airlinecode ? String(r.Airlinecode).trim() : null,
					Rpcid:        r.Rpcid,
					Iatacode:     r.Iatacode,
					Aircrafttype: r.Aircrafttype,
					Mtowquantity: r.Mtowquantity ? String(r.Mtowquantity).trim() : null,
					Mtowunit:     r.Mtowunit,
					IsActiveEntity: true,
					// Mtowcategory: r.Mtowcategory ? String(r.Mtowcategory).trim() : null,
					// Localcreatedby: "",
					// Localcreatedat: now,
					// Locallastchangedat: now,
					// Lastchangedat: now,
					// ⚠️ Do NOT include ExcelRowNumber if it's not part of your metadata
					};

					// Create transient entity
					const ctx = oListBinding.create(payload);

					// try {
					// 	// Wait for server round-trip
					// 	await ctx.created();

					// 	// Now safely read from the same context
					// 	const obj = typeof ctx.getObject === "function" ? ctx.getObject() : null;

					// 	mm.addMessages(new sap.ui.core.message.Message({
					// 	message: "Created: " + (obj?.AIRLINECODE || `(row ${i + 1})`),
					// 	persistent: true,
					// 	type: sap.ui.core.MessageType.Success
					// 	}));
					// } catch (e) {
					// 	// Creation failed for this row; add message and stop the whole run
					// 	mm.addMessages(new sap.ui.core.message.Message({
					// 	message: `Row ${i + 1} failed: ${e?.message || e}`,
					// 	persistent: true,
					// 	type: sap.ui.core.MessageType.Error
					// 	}));
					// 	throw e;
					// }
				}

				if (typeof fnResolve === "function") fnResolve();
			} catch (err) {
				if (typeof fnReject === "function") fnReject(err);
			}
		},
		getLocalISOString: function () {
			const d = new Date();
			const tzOffset = d.getTimezoneOffset() * 60000;
			const localISOTime = new Date(d - tzOffset).toISOString().slice(0, -1);
			return localISOTime;
		},
		_getUploadSet: function () {
		return this.byId("uploadSet") ||
				sap.ui.getCore().byId("excel_upload--uploadSet") ||
				sap.ui.core.Fragment.byId("excel_upload", "uploadSet");
		}		
	});
});
