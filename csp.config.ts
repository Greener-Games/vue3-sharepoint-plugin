export const cspPolicy = {
  "default-src": ["'self'"],
  "script-src": ["'self'", "'unsafe-inline'", "'unsafe-eval'"],
  "style-src": ["'self'", "'unsafe-inline'"],
  "img-src": ["'self'", "data:", "blob:"],
  "font-src": ["'self'", "data:"],
  "connect-src": ["'self'"],
};

export const cspString = Object.entries(cspPolicy)
  .map(([key, values]) => `${key} ${values.join(" ")}`)
  .join("; ");
