module.exports = function(grunt) {
    'use strict';

    grunt.initConfig({
        pkg: grunt.file.readJSON('package.json'),
	coveralls:{
	    test:{
		src:'coverage/lcov.info'
	    }
	}
    });
    grunt.loadNpmTasks('grunt-coveralls');
    grunt.registerTask('default', ['coveralls']);
};
