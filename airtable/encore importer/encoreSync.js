// Airtable Script: Excel Data Sync for All Searches

// Function to calculate codeClean based on the specified formula
function calculatecodeClean(code) {
	if (!code || typeof code !== "string") return null;

	const cleaned = code.trim();

	// Accept "AUTO" exactly (case-insensitive)
	if (/^AUTO$/i.test(cleaned)) {
		return "AUTO";
	}

	// Match pattern: ####, ####A, ####-A (ignores anything after)
	const match = cleaned.match(/^(\d+)-?([A-Za-z]){0,1}( |$)/);

	if (match) {
		const number = match[1];
		const letter = match[2] || "";
		return number + letter.toUpperCase(); // Always return uppercase for consistency
	}

	// Fallback
	return null;
}

function normalizeValue(value) {
    if (value instanceof Date) {
        return value.toISOString().split("T")[0];
    } else if (typeof value === "string") {
        // Normalize newlines and trim
        const normalized = value.replace(/\r\n/g, '\n').trim();

        // If it looks like a date, try to parse it
        const parsedDate = new Date(normalized);
        if (!isNaN(parsedDate.getTime())) {
            return parsedDate.toISOString().split("T")[0];
        }

        return normalized;
    }

    if (typeof value === "object" && value !== null && "name" in value) {
        return value.name;
    }

    return value;
}

function generateUpdateLog(existingRecord, newData, codeClean) {
	const changedFields = {};
	const previousValues = [];

	const fieldsToCompare = [...Object.keys(newData), "codeClean"];

	for (const field of fieldsToCompare) {
		const oldVal = field === "codeClean" 
			? existingRecord.getCellValue("codeClean") 
			: existingRecord.getCellValue(field);

		const newVal = field === "codeClean"
			? codeClean
			: newData[field];

		const normalizedOld = normalizeValue(oldVal);
		const normalizedNew = normalizeValue(newVal);

		if (normalizedOld !== normalizedNew) {
			changedFields[field] = newVal;
			previousValues.push(`**${field}:**\n${normalizedOld ?? "_None_"}`);
		}
	}

	if (Object.keys(changedFields).length === 0) {
		return null; // No changes
	}

	const now = new Date();
	const header = `### Updated ${now.toISOString().split("T")[0]} at ${now.toTimeString().slice(0,5)}`;
	const subheader = `#### Previous field values`;
	const logEntry = `${header}\n${subheader}\n\n${previousValues.join("\n\n")}`;

	const existingLog = existingRecord.getCellValue("updateLog") || "";
	const newLog = `${existingLog}\n\n${logEntry}`.trim();

	return {
		changedFields,
		newLog
	};
}

// Main function wrapped in try/catch for better error handling
async function runSync() {
	try {

		// Fetch existing records to check for matches
		// output.text(`fetching existing records to check for matches`);
		let existingRecords = await table.selectRecordsAsync(); // fetch all fields

		// Create a map of existing records by codeClean
		// output.text(`creating map of existing records`);
		let existingRecordsMap = {};
		existingRecords.records.forEach(record => {
			const codeClean = record.getCellValue('codeClean');
			if (codeClean !== null && codeClean !== undefined) {
				existingRecordsMap[codeClean] = record;
			}
		});

		// Prep update/create arrays
		let recordsToUpdate = [];
		let recordsToCreate = [];

		for (const row of rows) {
			const codeClean = calculatecodeClean(row.Code); // Calculate codeClean
			// output.text(`code: ${row.Code} >> codeClean: ${codeClean}`) // log progress
			
			if (codeClean === null) {
				console.log(`Skipping row with invalid Code: ${JSON.stringify(row)}`);
				continue;
			}
			
			let fields = { ...row }; // Clone all fields
			fields.codeClean = codeClean; // Add new field
			
			const existingRecord = existingRecordsMap[codeClean]; // Check if record exists

			if (existingRecord) {
				const result = generateUpdateLog(existingRecord, row, codeClean);
				if (result) {
					recordsToUpdate.push({
						id: existingRecord.id,
						fields: {
							...result.changedFields,
							updateLog: result.newLog
						}
					});
				}
			} else {
				// New record — prepare full fieldset
				recordsToCreate.push({ fields });
			}
		}

		// Summary counters
		const updateCount = recordsToUpdate.length;
		const createCount = recordsToCreate.length;

		// Process updates in batches of 50
		while (recordsToUpdate.length > 0) {
			output.text(`🔁 Updating ${recordsToUpdate.length} records...`);
			console.log("Update batch preview:", JSON.stringify(recordsToUpdate.slice(0, 5), null, 2));
			await table.updateRecordsAsync(recordsToUpdate.slice(0, 50));
			recordsToUpdate = recordsToUpdate.slice(50);
		}

		// Process creates in batches of 50
		if (recordsToCreate.length > 0) {
			output.text(`➕ Creating ${recordsToCreate.length} new records...`);
			await table.createRecordsAsync(recordsToCreate.slice(0, 50));
			recordsToCreate = recordsToCreate.slice(50);
		}

		// Final summary
		output.text(`✅ Sync complete:
		- Updated records: ${updateCount}
		- Created records: ${createCount}
		- Total processed: ${rows.length}`);
		
	} catch (error) {
		// Display error details for debugging
		console.error('Error during sync:', error);
		output.text(`Error: ${error.message || 'Unknown error occurred'}`);
	}
}

// Top-level file input prompt
let fileResult = await input.fileAsync(
	'Upload Excel or CSV file',
	{
		allowedFileTypes: ['.csv', '.xlsx', '.xls'], 
		hasHeaderRow: true
	}
);

let table = base.getTable('searches');
let rows;

// Determine file structure and extract rows
if (Array.isArray(fileResult.parsedContents)) {
	rows = fileResult.parsedContents; // CSV format
} else if (typeof fileResult.parsedContents === 'object' && fileResult.parsedContents !== null) {
	// Excel: get first sheet's rows
	const sheetNames = Object.keys(fileResult.parsedContents);
	if (sheetNames.length === 0) {
		throw new Error('Excel file has no sheets.');
	}
	rows = fileResult.parsedContents[sheetNames[0]];
} else {
	throw new Error('Unsupported or invalid file format.');
}

// Validate rows
if (!rows || !Array.isArray(rows) || rows.length === 0) {
	throw new Error('No valid data rows found in file.');
}

// Check if the required Code field exists
if (rows[0] && !rows[0].hasOwnProperty('Code')) {
	throw new Error('Required field "Code" is missing in the uploaded file.');
}

// Confirm import
let shouldContinue = await input.buttonsAsync(
	`Import ${rows.length} records from ${fileResult.file.name}?`,
	[{label: 'Yes', variant: 'primary'}, 'No']
);

if (shouldContinue !== 'Yes') {
	output.text('Import cancelled.');
} else {
	output.text('Processing data, please wait...');
	await runSync();
	output.text('✅ Sync complete.');
}