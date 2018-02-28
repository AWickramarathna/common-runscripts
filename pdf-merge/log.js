var Log = function(filename) {
	this.fso = new ActiveXObject('Scripting.FileSystemObject');
	this.file = this.fso.OpenTextFile(filename,
	                                  2,    // for writing (1=read, 2=write, 8=append)
	                                  true, // can create (false=cannot create)
	                                  0);   // unicode (0=ASCII, -1=unicode, -2=system default)
};

Log.prototype = {

	pad2: function (n) {
		return (n < 10) ? '0' + n : n;
	},

	pad3: function (n) {
		if (n < 10)
			return '00' + n;
		if (n < 100)
			return '0' + n;
		return n;
	},

	log: function (msg) {
		var d = new Date();
		var line = d.getUTCFullYear() + '-' +
			this.pad2(d.getUTCMonth() + 1) + '-' +
			this.pad2(d.getUTCDate()) + 'T' +
			this.pad2(d.getUTCHours()) + ':' +
			this.pad2(d.getUTCMinutes()) + ':' +
			this.pad2(d.getUTCSeconds()) + '.' +
			this.pad3(d.getUTCMilliseconds()) + 'Z ' +
			msg;
		this.file.WriteLine(line);
	},

	stat: function (msg) {
		this.log('STAT ' + msg);
	},

	prog: function (msg) {
		this.log('PROG ' + msg);
	},

	debug: function (msg) {
		this.log('DBUG ' + msg);
	},

	info: function (msg) {
		this.log('INFO ' + msg);
	},

	warn: function (msg) {
		this.log('WARN ' + msg);
	},

	error: function (msg) {
		this.log('ERRO ' + msg);
	},

	commenced: function () {
		this.stat('COMMENCED');
		this.info('Job commenced.');
	},

	completed: function () {
		this.info('Job completed.');
		this.stat('COMPLETED');
	},

	cancelled: function () {
		this.info('Job cancelled.');
		this.stat('CANCELLED');
	},

	close: function () {
		if (this.file != null) {
			this.file.Close();
		}
	}

};