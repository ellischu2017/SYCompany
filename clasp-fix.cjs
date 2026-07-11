const https = require('https');
const origAgent = https.Agent;
class PatchedAgent extends origAgent {
  constructor(options) {
    super({ ...options, rejectUnauthorized: false });
  }
}
https.Agent = PatchedAgent;
https.globalAgent = new PatchedAgent();
