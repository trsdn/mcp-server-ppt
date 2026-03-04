import * as vscode from 'vscode';
import * as path from 'path';

/**
 * PptMcp VS Code Extension
 *
 * This extension provides MCP server definitions for the PptMcp MCP server,
 * enabling AI assistants like GitHub Copilot to interact with Microsoft PowerPoint
 * through native COM automation.
 *
 * The extension bundles self-contained executables for both the MCP server and CLI -
 * no .NET SDK or runtime installation required.
 *
 * Agent Skills are registered via the chatSkills contribution point in package.json.
 */

export async function activate(context: vscode.ExtensionContext) {
	console.log('PptMcp extension is now active');

	// Register MCP server definition provider
	context.subscriptions.push(
		vscode.lm.registerMcpServerDefinitionProvider('ppt-mcp', {
			provideMcpServerDefinitions: async () => {
				// Return the MCP server definition for PptMcp
				const extensionPath = context.extensionPath;
				const mcpServerPath = path.join(extensionPath, 'bin', 'PptMcp.McpServer.exe');

				return [
					new vscode.McpStdioServerDefinition(
						'ppt-mcp',
						mcpServerPath,
						[],
						{
							// Optional environment variables can be added here if needed
						}
					)
				];
			}
		})
	);

	// Show welcome message on first activation
	const hasShownWelcome = context.globalState.get<boolean>('pptmcp.hasShownWelcome', false);
	if (!hasShownWelcome) {
		showWelcomeMessage();
		context.globalState.update('pptmcp.hasShownWelcome', true);
	}
}

function showWelcomeMessage() {
	const message = 'PptMcp extension activated! The PowerPoint MCP server is now available for AI assistants.';
	const learnMore = 'Learn More';

	vscode.window.showInformationMessage(message, learnMore).then(selection => {
		if (selection === learnMore) {
			vscode.env.openExternal(vscode.Uri.parse('https://github.com/sbroenne/mcp-server-ppt'));
		}
	});
}

export function deactivate() {
	console.log('PptMcp extension is now deactivated');
}
