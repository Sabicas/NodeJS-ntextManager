let ntext = require('../ntextManager.js');
let assert = require('assert');

describe('createDirVars',function(){
	it('should return undefined when client name is not found', function(){
		assert.equal(undefined,ntext.createDirVars('notarealclient'));
	});
});