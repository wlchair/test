module.exports = function(grunt) {
    'use strict';
    grunt.initConfig({
        pkg: grunt.file.readJSON('package.json'),
        jasmine: {
            pivotal: {
                src: 'src/out.js',
                options: {
                    specs: 'test/out-test.js',
                    keepRunner: true
                }
            }
        }
    });
    grunt.loadNpmTasks('grunt-contrib-jasmine');
    grunt.registerTask('test', ['jasmine']);
    grunt.registerTask('default', ['test']);
};