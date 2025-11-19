import Redis from "ioredis";
import "dotenv/config";

const redisUrl = process.env.REDIS_URL;

// Parse the Redis URL to check if it's Upstash (requires TLS)
const isUpstash = redisUrl.includes('upstash.io');

const redis = new Redis(redisUrl, {
  maxRetriesPerRequest: 3,
  enableReadyCheck: false,
  tls: isUpstash ? {} : undefined, // Enable TLS for Upstash
  family: 0, // Use IPv4 and IPv6
  retryStrategy(times) {
    if (times > 3) {
      console.error('✗ Redis max retries reached, giving up');
      return null; // Stop retrying
    }
    const delay = Math.min(times * 50, 2000);
    return delay;
  }
});

// Handle connection events
redis.on('connect', () => {
  console.log('✓ Redis connected successfully');
});

redis.on('ready', () => {
  console.log('✓ Redis ready to accept commands');
});

redis.on('error', (err) => {
  console.error('✗ Redis error:', err.message);
  // Don't crash the app on Redis errors
});

redis.on('close', () => {
  console.log('⚠ Redis connection closed');
});

redis.on('reconnecting', () => {
  console.log('⟳ Redis reconnecting...');
});

export default redis;