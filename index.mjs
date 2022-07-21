import Fastify from "fastify";
import * as fs from "fs";
import { readFile, set_fs, utils } from "xlsx/xlsx.mjs";
// import excelDateToJSDate from "./utils/excelDateToJSDate.mjs";

set_fs(fs);

const START_COLUMN = "F";
const END_COLUMN = "AH";
const START_ROW = 58;
const END_ROW = 159;

const workbook = readFile("access-list.xlsx");

const fastify = Fastify({
	logger: true,
});

fastify.get("/", async () => {
	const sheet = workbook.Sheets["Access List"];

	// get all the weeks
	sheet["!ref"] = `A${START_ROW}:A${END_ROW}`;
	const allWeeks = utils.sheet_to_json(sheet, { header: 1, raw: false }).flat();

	// get the classes
	sheet["!ref"] = `${START_COLUMN}1:${END_COLUMN}1`;
	const classesArray = utils.sheet_to_json(sheet, { header: "A" });
	const classes = classesArray[0];

	const classesToExport = [];

	const finalArray = {};

	// use this method in order to be able to break early
	const keys = Object.keys(classes);

	for (let i = 0; i < keys.length; i++) {
		if (i !== 3) continue;

		const column = keys[i];
		const studentClass = classes[column];
		classesToExport.push(studentClass);

		const range = `${column}${START_ROW}:${column}${END_ROW}`;

		sheet["!ref"] = range;
		const classWeekContent = utils.sheet_to_json(sheet, { header: 1 });

		const classContent = [];

		classWeekContent.forEach((weekContent, i) => {
			const week = allWeeks[i];
			classContent.push({ [week]: weekContent.toString() });
		});

		finalArray[studentClass] = classContent;
	}

	return { weeks: allWeeks, classes: classesToExport, data: finalArray };
});

const start = async () => {
	try {
		await fastify.listen({ port: 3000 });
	} catch (err) {
		fastify.log.error(err);
		process.exit(1);
	}
};
start();
