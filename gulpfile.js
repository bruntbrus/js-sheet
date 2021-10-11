/*
 * Gulp tasks
 */

/* jshint node: true */

const gulp = require('gulp')

// Configuration
const config = {
  srcDir: 'src',
}

// Runs lint task by default
gulp.task('default', ['lint'])

// Runs all lint tasks except for node scripts
gulp.task('lint', ['lint:src'])

// Lint checks scripts
gulp.task('lint:src', () => lint(config.srcDir + '/*.js'))

// Lint checks this file
gulp.task('lint:gulpfile', () => lint('gulpfile.js'))

// Lint checks files
function lint(globs) {
  const sort = require('gulp-sort')
  const jshint = require('gulp-jshint')

  return gulp.src(globs)
    .pipe(sort())
    .pipe(jshint())
    .pipe(jshint.reporter('jshint-stylish'))
}
