module.exports = function(grunt) {
    'use strict';

    grunt.initConfig({
        pkg: grunt.file.readJSON('package.json')
    });


    // Default task(s).
    // grunt.registerTask('default', ['jsbint:all', 'dist']);
    // grunt.registerTask('dist', ['build', 'uglify', 'copy']);
    // grunt.registerTask('deploy', ['doc', 'jekyll', 'gh-pages']);
    // grunt.registerTask('test', ['connect', 'qunit']);
};
