import { createProgram } from "./cli.js";

const program = createProgram();
program.parseAsync(process.argv);
