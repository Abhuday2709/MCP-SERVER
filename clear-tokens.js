// Script to clear all tokens from Redis
import redis from './config/redisClient.js';

console.log('Clearing all user tokens from Redis...\n');

try {
  // Get all keys matching user:token:*
  const keys = await redis.keys('user:token:*');
  
  if (keys.length === 0) {
    console.log('No tokens found in Redis.');
  } else {
    console.log(`Found ${keys.length} token(s). Deleting...`);
    
    for (const key of keys) {
      await redis.del(key);
      console.log(`✓ Deleted: ${key}`);
    }
    
    console.log(`\n✓ Successfully cleared ${keys.length} token(s)`);
  }
  
  await redis.quit();
  console.log('\nDone! Now:');
  console.log('1. Restart your backend server');
  console.log('2. Go to your app and click "Connect" on Teams');
  console.log('3. Sign in again to get fresh tokens with new permissions');
} catch (error) {
  console.error('Error:', error.message);
  process.exit(1);
}
