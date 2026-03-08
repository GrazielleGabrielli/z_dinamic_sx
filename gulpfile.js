'use strict';

const build = require('@microsoft/sp-build-web');
const postcss = require('gulp-postcss');
const tailwind = require('tailwindcss');
const autoprefixer = require('autoprefixer');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

const tailwindcss = build.subTask('tailwindcss', function (gulp, buildOptions, done) {
  return gulp
    .src('src/assets/tailwind.css')
    .pipe(postcss([tailwind('./tailwind.config.js'), autoprefixer()]))
    .pipe(gulp.dest('src/assets/dist'))
    .on('end', done);
});

build.rig.addPreBuildTask(tailwindcss);

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

build.initialize(require('gulp'));
