import crypto from 'crypto';

/**
 * Authentication middleware
 * Accepts two authentication methods:
 * 1. x-andcon-secret header (server-to-server)
 * 2. x-andcon-token header (custom HMAC-SHA256 token for browser uploads)
 */
export function requireAuth(req, res, next) {
  const secret = req.headers['x-andcon-secret'];
  const token = req.headers['x-andcon-token'];

  if (!process.env.PROCESSOR_SHARED_SECRET) {
    return res.status(500).json({ error: 'Server configuration error: PROCESSOR_SHARED_SECRET not set' });
  }

  // Option 1: Direct shared secret (for server-to-server calls)
  if (secret === process.env.PROCESSOR_SHARED_SECRET) {
    return next();
  }

  // Option 2: Custom HMAC-SHA256 token (for direct client uploads)
  if (token) {
    try {
      // Token format: base64(payload).signature
      const [payloadB64, signature] = token.split('.');

      if (!payloadB64 || !signature) {
        return res.status(401).json({ error: 'Unauthorized: Invalid token format' });
      }

      // Decode payload
      const payloadString = Buffer.from(payloadB64, 'base64').toString('utf-8');
      const payload = JSON.parse(payloadString);

      // Verify signature
      const expectedSignature = crypto
        .createHmac('sha256', process.env.PROCESSOR_SHARED_SECRET)
        .update(payloadString)
        .digest('hex');

      if (signature !== expectedSignature) {
        return res.status(401).json({ error: 'Unauthorized: Invalid token signature' });
      }

      // Check expiry
      const now = Math.floor(Date.now() / 1000);
      if (!payload.exp || payload.exp < now) {
        return res.status(401).json({ error: 'Unauthorized: Token expired' });
      }

      // Token is valid
      req.tokenData = payload;
      return next();

    } catch (err) {
      console.error('Token validation error:', err.message);
      return res.status(401).json({ error: 'Unauthorized: Token validation failed' });
    }
  }

  // Neither secret nor valid token provided
  return res.status(401).json({ error: 'Unauthorized: Missing or invalid authentication' });
}
