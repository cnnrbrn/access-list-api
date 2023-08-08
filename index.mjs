import Fastify from "fastify";
import * as fs from "fs";
import { readFile, set_fs, utils } from "xlsx/xlsx.mjs";

set_fs(fs);

const START_COLUMN = "I";
const END_COLUMN = "AM";
const START_ROW = 33;
const END_ROW = 231;
const WEEK_COLUMN = "B";

const workbook = readFile("access-list.xlsx");

const fastify = Fastify({
	logger: true,
});

fastify.get("/", async () => {
	const sheet = workbook.Sheets["Access List"];

	// get all the weeks
	sheet["!ref"] = `A${START_ROW}:${WEEK_COLUMN}${END_ROW}`;
	const allWeeks = utils.sheet_to_json(sheet, { header: 1, raw: false }).flat();
	console.log("allWeeks", allWeeks);

	// get the classes
	sheet["!ref"] = `${START_COLUMN}1:${END_COLUMN}1`;
	const classesArray = utils.sheet_to_json(sheet, { header: "A" });
	const classes = classesArray[0];

	const classesToExport = [];

	const finalArray = {};

	const keys = Object.keys(classes);

	const classesToSkip = ["Aug 19 F", "Mar 20 F", "Aug 20 F", "FED2 Oct 21 F", "Jan 21 F"];

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
