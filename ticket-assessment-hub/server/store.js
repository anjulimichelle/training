'use strict';
// JSON-file store for the staging prototype. Writes are atomic (temp file + rename)
// and serialized, so a crash mid-write never corrupts raw submissions.
const fs = require('fs');
const path = require('path');

const EMPTY = () => ({ meta: { createdAt: new Date().toISOString(), mode: 'staging' }, trainees: [], quizzes: [], submissions: [], evaluations: [], auditLog: [] });

class Store {
  constructor(file) {
    this.file = file;
    fs.mkdirSync(path.dirname(file), { recursive: true });
    if (fs.existsSync(file)) {
      this.db = JSON.parse(fs.readFileSync(file, 'utf8'));
      for (const k of Object.keys(EMPTY())) if (!this.db[k]) this.db[k] = EMPTY()[k];
    } else {
      this.db = EMPTY();
      this.flush();
    }
    this.queue = Promise.resolve();
  }

  flush() {
    const tmp = this.file + '.tmp';
    fs.writeFileSync(tmp, JSON.stringify(this.db, null, 2));
    fs.renameSync(tmp, this.file);
  }

  // Run a mutation and persist it. Mutations run one at a time.
  mutate(fn) {
    const run = this.queue.then(() => {
      const out = fn(this.db);
      this.flush();
      return out;
    });
    this.queue = run.catch(() => {});
    return run;
  }

  audit(db, action, details) {
    db.auditLog.push({ at: new Date().toISOString(), action, ...details });
    if (db.auditLog.length > 5000) db.auditLog.splice(0, db.auditLog.length - 5000);
  }

  reset() {
    return this.mutate((db) => {
      Object.assign(db, EMPTY());
    });
  }
}

module.exports = { Store };
