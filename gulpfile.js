'use strict';

const build = require('@microsoft/sp-build-web');

// Suppress warnings for the local CSS class 'ms-Grid' not being camelCase
build.addSuppression(/Warning - \[sass\] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe\./);

// Get the default set of Gulp tasks
var getTasks = build.rig.getTasks;

build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  // Comment or remove this line to use the default 'serve' task
  result.set('serve', result.get('serve-deprecated'));

  return result;
};

// Initialize the build process
/* fast-serve */
const { addFastServe } = require("spfx-fast-serve-helpers");
addFastServe(build);
/* end of fast-serve */

build.initialize(require('gulp'));

