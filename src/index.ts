import chalk from "chalk";
import { createProgram } from "./cli.js";

const program = createProgram();
program.parseAsync(process.argv).catch((e: unknown) => {
  console.error(chalk.red((e as Error).message ?? String(e)));
  process.exit(1);
});
