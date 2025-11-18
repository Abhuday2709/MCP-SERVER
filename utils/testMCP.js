import mcpClient from '../mcp/client/mcp-client.js';

async function testMCP() {
  try {
    console.log('=== Testing Gmail MCP Server ===\n');
    
    // Test 1: Get available tools
    console.log('Test 1: Fetching available Gmail tools...');
    const gmailTools = await mcpClient.getAvailableTools(['gmail']);
    console.log('\nAvailable Gmail tools:');
    gmailTools.forEach(tool => {
      console.log(`  ✓ ${tool.name}`);
      console.log(`    Description: ${tool.description}`);
      console.log(`    Input Schema:`, JSON.stringify(tool.inputSchema, null, 2));
      console.log('');
    });

    console.log(`\n✅ Successfully loaded ${gmailTools.length} Gmail tools\n`);
    
    // Clean up
    await mcpClient.disconnect();
    console.log('Cleanup completed');
    process.exit(0);
  } catch (error) {
    console.error('\n❌ Test failed:', error);
    await mcpClient.disconnect();
    process.exit(1);
  }
}

// Add timeout to prevent hanging
setTimeout(() => {
  console.error('\n❌ Test timeout after 30 seconds');
  process.exit(1);
}, 30000);

testMCP();