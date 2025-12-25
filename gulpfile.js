const gulp = require('gulp');
const path = require('path');

gulp.task('build:icons', function() {
	// 建議直接使用相對路徑字串，這對 Gulp 的通配符支援最穩
	// 如果一定要用 path.resolve，通配符部分 '**/*.svg' 應該留在字串中
	const nodeSource = 'nodes/**/*.svg'; 
	const nodeDestination = 'dist/nodes';

	console.log(`正在複製圖示從 ${nodeSource} 到 ${nodeDestination}...`);

	return gulp.src(nodeSource)
		.pipe(gulp.dest(nodeDestination));
});