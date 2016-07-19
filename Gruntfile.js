
module.exports = function(grunt) {
    'use strict';
    grunt.initConfig({
        pkg: grunt.file.readJSON('package.json'),
	conventionalChangelog: {
    options: {
      changelogOpts: {
        // conventional-changelog options go here
        preset: 'jshint'
      },
      context: {
        // context goes here
      },
      gitRawCommitsOpts: {
        // git-raw-commits options go here
      },
      parserOpts: {
        // conventional-commits-parser options go here
      },
      writerOpts: {
        // conventional-changelog-writer options go here
      }
    },
    release: {
      src: 'CHANGELOG.md'
    }
  }
});

grunt.loadNpmTasks('grunt-conventional-changelog');
grunt.registerTask('default', ['conventionalChangelog']);
};
