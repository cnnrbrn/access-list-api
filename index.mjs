import Fastify from "fastify";
import cors from "@fastify/cors";
import * as fs from "fs";
import { readFile, set_fs, utils } from "xlsx/xlsx.mjs";

set_fs(fs);

const START_COLUMN = "C";
const END_COLUMN = "Z";
const START_ROW = 2;
const END_ROW = 168;
const WEEK_COLUMN = "A";

const workbook = readFile("access-list.xlsx");

const fastify = Fastify({
	logger: true,
});

await fastify.register(cors, {
	// put your options here
});

// Function to prepend "20" to a 2-digit year in a date string
const ensureFourDigitYear = (dateStr) => {
	if (dateStr && typeof dateStr === "string") {
		// Split the date string into parts
		const parts = dateStr.split("/");
		// Check if the year part (last part) is 2 digits
		if (parts.length === 3 && parts[2].length === 2) {
			parts[2] = "20" + parts[2];
			return parts.join("/");
		}
	}
	return dateStr;
};

fastify.get("/", async () => {
	const sheet = workbook.Sheets["Sheet1"];

	// Get all the weeks
	sheet["!ref"] = `A${START_ROW}:${WEEK_COLUMN}${END_ROW}`;
	const allWeeksRaw = utils.sheet_to_json(sheet, { header: 1, raw: false }).flat();

	// Ensure all dates in allWeeks have 4-digit years
	const allWeeks = allWeeksRaw.map(ensureFourDigitYear);
	console.log("allWeeks...", allWeeks);

	// Get the classes
	sheet["!ref"] = `${START_COLUMN}1:${END_COLUMN}1`;
	const classesArray = utils.sheet_to_json(sheet, { header: "A" });
	const classes = classesArray[0];

	const classesToExport = [];
	const finalArray = {};
	const keys = Object.keys(classes);

	const classesToSkip = ["Aug 19 F", "Mar 20 F", "Aug 20 F", "FED2 Oct 21 F", "Jan 21 F", "JAN20 P", "JAN22 F"];

	for (let i = 0; i < keys.length; i++) {
		const column = keys[i];
		const studentClass = classes[column];

		if (classesToSkip.includes(studentClass)) {
			continue;
		}

		classesToExport.push(studentClass);

		const range = `${column}${START_ROW}:${column}${END_ROW}`;
		sheet["!ref"] = range;
		const classWeekContent = utils.sheet_to_json(sheet, { header: 1 });

		const classContent = [];

		classWeekContent.forEach((weekContent, i) => {
			const week = ensureFourDigitYear(allWeeks[i]);
			classContent.push({ [week]: weekContent.toString() });
		});

		finalArray[studentClass] = classContent;
	}

	return { data: finalArray };
});

const start = async () => {
	try {
		await fastify.listen({ port: 3010 });
	} catch (err) {
		fastify.log.error(err);
		process.exit(1);
	}
};
start();
