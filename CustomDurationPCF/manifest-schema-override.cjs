const schema = require("pcf-scripts/ManifestSchema.json");
const allowedTypes = schema.definitions?.dataType?.enum;

if (!Array.isArray(allowedTypes) || !allowedTypes.includes("Whole.None")) {
	throw new Error("CustomDurationPCF: the pcf-scripts manifest schema changed; review the Whole.Duration override.");
}

// Extend the validator's in-memory schema without changing the installed package.
if (!allowedTypes.includes("Whole.Duration")) {
	allowedTypes.push("Whole.Duration");
}
