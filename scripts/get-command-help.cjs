#!/usr/bin/env node
/**
 * get-command-help.cjs - Lazy-load command documentation
 * 
 * Query command schemas, examples, and tips from command-schemas.json
 * without loading the entire documentation into context.
 * 
 * Usage:
 *   node scripts/get-command-help.cjs <command>           # Full details for a command
 *   node scripts/get-command-help.cjs --list              # List all commands
 *   node scripts/get-command-help.cjs --category <name>   # Commands in a category
 *   node scripts/get-command-help.cjs --schema <command>  # Just the JSON schema
 *   node scripts/get-command-help.cjs --example <command> # Just the example
 *   node scripts/get-command-help.cjs --errors            # List all error codes
 *   node scripts/get-command-help.cjs --categories        # List all categories
 *   node scripts/get-command-help.cjs --search <term>     # Search commands by keyword
 */

const fs = require('fs');
const path = require('path');

// Load schemas
const schemasPath = path.join(__dirname, '..', 'docs', 'command-schemas.json');
let schemas;

try {
  schemas = JSON.parse(fs.readFileSync(schemasPath, 'utf8'));
} catch (e) {
  console.error(`Error loading schemas: ${e.message}`);
  console.error(`Expected file at: ${schemasPath}`);
  process.exit(1);
}

// Parse arguments
const args = process.argv.slice(2);

if (args.length === 0) {
  printUsage();
  process.exit(0);
}

// Handle flags
const flag = args[0];

switch (flag) {
  case '--help':
  case '-h':
    printUsage();
    break;
    
  case '--list':
  case '-l':
    listAllCommands();
    break;
    
  case '--categories':
  case '-c':
    listCategories();
    break;
    
  case '--category':
    if (!args[1]) {
      console.error('Error: --category requires a category name');
      console.error('Use --categories to see available categories');
      process.exit(1);
    }
    showCategory(args[1]);
    break;
    
  case '--schema':
  case '-s':
    if (!args[1]) {
      console.error('Error: --schema requires a command name');
      process.exit(1);
    }
    showSchema(args[1]);
    break;
    
  case '--example':
  case '-e':
    if (!args[1]) {
      console.error('Error: --example requires a command name');
      process.exit(1);
    }
    showExample(args[1]);
    break;
    
  case '--errors':
    listErrors();
    break;
    
  case '--search':
    if (!args[1]) {
      console.error('Error: --search requires a search term');
      process.exit(1);
    }
    searchCommands(args[1]);
    break;
    
  default:
    // Assume it's a command name
    if (flag.startsWith('-')) {
      console.error(`Unknown flag: ${flag}`);
      printUsage();
      process.exit(1);
    }
    showCommandDetails(flag);
}

// === Functions ===

function printUsage() {
  console.log(`
Windows Desktop Automation - Command Help (v${schemas.version})
${'='.repeat(50)}

Usage:
  node get-command-help.cjs <command>           Full details for a command
  node get-command-help.cjs --list              List all commands
  node get-command-help.cjs --categories        List all categories
  node get-command-help.cjs --category <name>   Commands in a category
  node get-command-help.cjs --schema <command>  Just the JSON schema
  node get-command-help.cjs --example <command> Just the example
  node get-command-help.cjs --errors            List all error codes
  node get-command-help.cjs --search <term>     Search commands by keyword

Examples:
  node get-command-help.cjs click
  node get-command-help.cjs --category mouse
  node get-command-help.cjs --schema explore_window
  node get-command-help.cjs --search window
`);
}

function listAllCommands() {
  console.log(`\nAll Commands (${Object.keys(schemas.commands).length} total):`);
  console.log('='.repeat(60));
  
  // Group by category
  for (const [catId, cat] of Object.entries(schemas.categories)) {
    console.log(`\n${cat.name}:`);
    for (const cmd of cat.commands) {
      const cmdData = schemas.commands[cmd];
      const desc = cmdData ? cmdData.description.substring(0, 50) : 'N/A';
      console.log(`  ${cmd.padEnd(20)} ${desc}...`);
    }
  }
}

function listCategories() {
  console.log('\nAvailable Categories:');
  console.log('='.repeat(40));
  
  for (const [catId, cat] of Object.entries(schemas.categories)) {
    console.log(`\n${catId.padEnd(18)} ${cat.name}`);
    console.log(`${''.padEnd(18)} ${cat.description}`);
    console.log(`${''.padEnd(18)} Commands: ${cat.commands.length}`);
  }
}

function showCategory(categoryName) {
  const catId = categoryName.toLowerCase().replace(/\s+/g, '_');
  const cat = schemas.categories[catId];
  
  if (!cat) {
    console.error(`Category not found: ${categoryName}`);
    console.error('\nAvailable categories:');
    for (const id of Object.keys(schemas.categories)) {
      console.error(`  - ${id}`);
    }
    process.exit(1);
  }
  
  console.log(`\n${cat.name} Commands`);
  console.log('='.repeat(50));
  console.log(`${cat.description}\n`);
  
  for (const cmdName of cat.commands) {
    const cmd = schemas.commands[cmdName];
    if (cmd) {
      console.log(`${cmdName}`);
      console.log(`  ${cmd.description}`);
      console.log(`  Required: ${cmd.required.join(', ') || 'none'}`);
      console.log(`  Optional: ${cmd.optional.join(', ') || 'none'}`);
      console.log();
    }
  }
}

function showSchema(commandName) {
  const cmd = schemas.commands[commandName];
  
  if (!cmd) {
    console.error(`Command not found: ${commandName}`);
    suggestSimilar(commandName);
    process.exit(1);
  }
  
  console.log(`\nSchema for: ${commandName}`);
  console.log('='.repeat(40));
  console.log(JSON.stringify(cmd.schema, null, 2));
  console.log(`\nRequired: ${cmd.required.join(', ')}`);
  console.log(`Optional: ${cmd.optional.join(', ') || 'none'}`);
}

function showExample(commandName) {
  const cmd = schemas.commands[commandName];
  
  if (!cmd) {
    console.error(`Command not found: ${commandName}`);
    suggestSimilar(commandName);
    process.exit(1);
  }
  
  console.log(`\nExample for: ${commandName}`);
  console.log('='.repeat(40));
  console.log(JSON.stringify(cmd.example, null, 2));
}

function showCommandDetails(commandName) {
  const cmd = schemas.commands[commandName];
  
  if (!cmd) {
    console.error(`Command not found: ${commandName}`);
    suggestSimilar(commandName);
    process.exit(1);
  }
  
  console.log(`
${'='.repeat(60)}
Command: ${commandName}
Category: ${cmd.category}
${'='.repeat(60)}

${cmd.description}

REQUIRED PARAMETERS:
${formatParameters(cmd.schema, cmd.required)}

OPTIONAL PARAMETERS:
${cmd.optional.length > 0 ? formatParameters(cmd.schema, cmd.optional) : '  (none)'}

EXAMPLE:
${JSON.stringify(cmd.example, null, 2)}

RETURNS:
${JSON.stringify(cmd.returns, null, 2)}

TIPS:
${cmd.tips ? cmd.tips.map(t => `  - ${t}`).join('\n') : '  (none)'}
`);
}

function formatParameters(schema, params) {
  if (!params || params.length === 0) return '  (none)';
  
  return params.map(p => {
    const paramSchema = schema[p];
    if (!paramSchema) return `  - ${p}`;
    
    const type = paramSchema.type || 'any';
    const desc = paramSchema.description || '';
    const def = paramSchema.default !== undefined ? ` (default: ${paramSchema.default})` : '';
    const constVal = paramSchema.const ? ` = "${paramSchema.const}"` : '';
    
    return `  - ${p} (${type})${constVal}: ${desc}${def}`;
  }).join('\n');
}

function listErrors() {
  console.log('\nError Codes:');
  console.log('='.repeat(60));
  
  for (const [code, info] of Object.entries(schemas.error_codes)) {
    console.log(`\n${code}`);
    console.log(`  Meaning:  ${info.meaning}`);
    console.log(`  Recovery: ${info.recovery}`);
  }
}

function searchCommands(term) {
  const searchTerm = term.toLowerCase();
  const matches = [];
  
  for (const [name, cmd] of Object.entries(schemas.commands)) {
    const searchableText = `${name} ${cmd.description} ${cmd.category}`.toLowerCase();
    if (searchableText.includes(searchTerm)) {
      matches.push({ name, cmd });
    }
  }
  
  if (matches.length === 0) {
    console.log(`\nNo commands found matching: "${term}"`);
    return;
  }
  
  console.log(`\nCommands matching "${term}" (${matches.length} found):`);
  console.log('='.repeat(50));
  
  for (const { name, cmd } of matches) {
    console.log(`\n${name} [${cmd.category}]`);
    console.log(`  ${cmd.description}`);
  }
}

function suggestSimilar(commandName) {
  const allCommands = Object.keys(schemas.commands);
  const similar = allCommands.filter(cmd => 
    cmd.includes(commandName) || 
    commandName.includes(cmd) ||
    levenshteinDistance(cmd, commandName) <= 3
  );
  
  if (similar.length > 0) {
    console.error('\nDid you mean:');
    for (const s of similar.slice(0, 5)) {
      console.error(`  - ${s}`);
    }
  }
}

function levenshteinDistance(a, b) {
  const matrix = Array(b.length + 1).fill(null).map(() => Array(a.length + 1).fill(null));
  
  for (let i = 0; i <= a.length; i++) matrix[0][i] = i;
  for (let j = 0; j <= b.length; j++) matrix[j][0] = j;
  
  for (let j = 1; j <= b.length; j++) {
    for (let i = 1; i <= a.length; i++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      matrix[j][i] = Math.min(
        matrix[j][i - 1] + 1,
        matrix[j - 1][i] + 1,
        matrix[j - 1][i - 1] + cost
      );
    }
  }
  
  return matrix[b.length][a.length];
}
