module.exports = function(grunt) {
    'use strict';

	// Project configuration.
	grunt.initConfig({
		// Task configuration.
		jshint: {
            options: {
                jshintrc: '.jshintrc',
            },
			gruntfile: {
				src: 'Gruntfile.js'
			}
		},
        iisexpress: {
            server: {
                options: {
                    port: 3000,
                    killOn: 'aspunit.finish'
                }
            }
        },
        aspunit: {
            test: {
                options: {
                    urls: ['http://localhost:3000/test/'],
                    force: true
                }
            }
        }
	});

	// These plugins provide necessary tasks.
	grunt.loadNpmTasks('grunt-contrib-jshint');
    grunt.loadNpmTasks('grunt-iisexpress');
	grunt.loadNpmTasks('grunt-aspunit');

    grunt.registerTask('default', ['jshint']);
	grunt.registerTask('test', ['iisexpress', 'aspunit']);

};