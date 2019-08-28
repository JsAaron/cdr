var child_process = require('child_process');


// exec('python ./command/app.py test', function(error, stdout, stderr) {
//   if (error) {
//     console.error('error: ' + error);
//     return;
//   }
//   console.log('receive: ' + stdout);

// });

var workerProcess = child_process.spawn('python', ['./command/app.py', "test"]);
workerProcess.stdout.on('data', function(data) {
  console.log('消息: ' + data);
});
workerProcess.stderr.on('data', function(data) {
  console.log('错误: ' + data);
});
workerProcess.on('close', function(code) {
  console.log('子进程已退出，退出码 ' + code);
})