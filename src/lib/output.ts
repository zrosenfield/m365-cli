import { GraphError } from "./graph.js";

export function outputData(data: unknown, raw = false): void {
  if (raw) {
    process.stdout.write(JSON.stringify(data));
  } else {
    process.stdout.write(JSON.stringify({ data }, null, 2) + "\n");
  }
}

export function outputError(err: unknown, raw = false): void {
  let error: { code: string; message: string; status?: number };
  if (err instanceof GraphError) {
    error = { code: err.code, message: err.message, status: err.status };
  } else if (err instanceof Error) {
    error = { code: "Error", message: err.message };
  } else {
    error = { code: "UnknownError", message: String(err) };
  }

  const out = raw ? JSON.stringify({ error }) : JSON.stringify({ error }, null, 2) + "\n";
  process.stderr.write(out);
}

export function handleCommandError(err: unknown): never {
  outputError(err);
  process.exit(1);
}
