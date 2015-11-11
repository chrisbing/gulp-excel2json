# gulp-excel2json
> Excel (XLSX/XLS) to json.


## Usage
First, install `gulp-excel2json` as a development dependency:

```shell
> npm install --save-dev gulp-excel2json
```

Then, add it to your `gulpfile.js`:

```javascript
var excel2json = require('gulp-excel2json');

gulp.task('copy', function() {
    gulp.src('config/**.xlsx')
        .pipe(excel2json({
            headRow: 1,
            valueRowStart: 3,
            trace: true
        }))
        .pipe(gulp.dest('build'))
});
```


## API

### excel2json([options])

#### options.headRow
Type: `number`

Default: `1`

The row number of head. (Start from 1).

#### options.valueRowStart
Type: `number`

Default: `3`

The start row number of values. (Start from 1)

#### options.trace
Type: `Boolean`

Default: `false`

Whether to log each file path while convert success.

## License
MIT &copy; Chris
