// Lets `node --test tests/` work on Node versions that resolve a directory argument as a module.
// (`node --test tests/*.test.js` and `npm test` run the same files.)
require('./core.test.js');
require('./xlsx.test.js');
