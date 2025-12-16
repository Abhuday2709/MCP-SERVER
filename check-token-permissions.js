// Script to check what permissions are in the current access token
import redis from './config/redisClient.js';
import jwt from 'jsonwebtoken';

console.log('\n=== Checking Token Permissions ===\n');

try {
  // Get all user tokens from Redis
  const keys = await redis.keys('user:token:*');
  
  if (keys.length === 0) {
    console.log('❌ No tokens found in Redis.');
    console.log('\nYou need to authenticate first:');
    console.log('1. Go to your app (http://localhost:5173)');
    console.log('2. Click "Connect" on Teams card');
    console.log('3. Sign in with your Microsoft account');
    await redis.quit();
    process.exit(0);
  }
  
  console.log(`Found ${keys.length} token(s) in Redis\n`);
  
  for (const key of keys) {
    const data = await redis.get(key);
    const userData = JSON.parse(data);
    
    console.log(`\n📋 User: ${userData.user?.email || 'Unknown'}`);
    console.log(`   Provider: ${userData.user?.provider || 'Unknown'}`);
    console.log(`   User ID: ${key.replace('user:token:', '')}`);
    
    if (userData.microsoftTokens && userData.microsoftTokens.accessToken) {
      const accessToken = userData.microsoftTokens.accessToken;
      
      console.log('\n🔍 Decoding Microsoft Access Token...\n');
      
      // Decode the JWT token (without verification since we just want to read it)
      try {
        const decoded = jwt.decode(accessToken, { complete: true });
        
        if (!decoded) {
          console.log('   ❌ Could not decode token');
          continue;
        }
        
        const payload = decoded.payload;
        
        // Check for scopes/permissions
        const scopes = payload.scp || payload.scope || '';
        const roles = payload.roles || [];
        
        console.log('   📜 Granted Scopes (scp):');
        if (scopes) {
          const scopeList = scopes.split(' ');
          scopeList.forEach(scope => {
            console.log(`      ✓ ${scope}`);
          });
          
          // Check for required Teams permissions
          console.log('\n   🔍 Checking Required Permissions:');
          const requiredScopes = [
            'User.Read',
            'Chat.Read',
            'Chat.ReadWrite',
            'ChatMessage.Read',
            'ChatMessage.Send'
          ];
          
          requiredScopes.forEach(required => {
            const hasPermission = scopeList.some(s => s.includes(required));
            if (hasPermission) {
              console.log(`      ✅ ${required} - GRANTED`);
            } else {
              console.log(`      ❌ ${required} - MISSING`);
            }
          });
          
          // Check for admin-required permissions
          console.log('\n   ⚠️  Admin-Required Permissions (should NOT be present):');
          const adminScopes = [
            'Team.ReadBasic.All',
            'Channel.ReadBasic.All',
            'ChannelMessage.Read.All',
            'ChannelMessage.Send'
          ];
          
          adminScopes.forEach(admin => {
            const hasPermission = scopeList.some(s => s.includes(admin));
            if (hasPermission) {
              console.log(`      ⚠️  ${admin} - PRESENT (requires admin consent)`);
            } else {
              console.log(`      ✓ ${admin} - Not present (good)`);
            }
          });
        } else {
          console.log('      ❌ No scopes found in token');
        }
        
        if (roles.length > 0) {
          console.log('\n   📜 Roles:');
          roles.forEach(role => {
            console.log(`      - ${role}`);
          });
        }
        
        // Token expiration
        if (payload.exp) {
          const expiresAt = new Date(payload.exp * 1000);
          const now = new Date();
          const isExpired = expiresAt < now;
          
          console.log('\n   ⏰ Token Expiration:');
          console.log(`      Expires: ${expiresAt.toLocaleString()}`);
          console.log(`      Status: ${isExpired ? '❌ EXPIRED' : '✅ Valid'}`);
          
          if (!isExpired) {
            const timeLeft = Math.floor((expiresAt - now) / 1000 / 60);
            console.log(`      Time left: ${timeLeft} minutes`);
          }
        }
        
        // Audience
        if (payload.aud) {
          console.log('\n   🎯 Audience:');
          console.log(`      ${payload.aud}`);
        }
        
        // Issuer
        if (payload.iss) {
          console.log('\n   🏢 Issuer:');
          console.log(`      ${payload.iss}`);
        }
        
      } catch (decodeError) {
        console.log(`   ❌ Error decoding token: ${decodeError.message}`);
      }
    } else {
      console.log('   ⚠️  No Microsoft access token found');
    }
    
    console.log('\n' + '='.repeat(70));
  }
  
  console.log('\n📊 Summary:\n');
  console.log('If you see "MISSING" permissions above:');
  console.log('1. Update Azure Portal to add those permissions');
  console.log('2. Run: node clear-tokens.js');
  console.log('3. Re-authenticate in your app');
  console.log('4. Run this script again to verify');
  
  console.log('\nIf you see "GRANTED" for all required permissions but still get errors:');
  console.log('1. Your organization may require admin consent for ALL Teams permissions');
  console.log('2. Contact your IT admin to grant consent');
  console.log('3. Or use a personal Microsoft account for testing');
  
  await redis.quit();
  
} catch (error) {
  console.error('❌ Error:', error.message);
  process.exit(1);
}
