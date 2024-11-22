var synth = window.speechSynthesis;
var spkText = ""; //需要阅读的内容
var voices = synth.getVoices();//获取音色
document.body.parentNode.style.overflow = "hidden";
window.resizeTo(200, 100);

//启动则开始循环读取文件
//发现文件不为空，则阅读一次。

function speak(timbre) {
	synth.cancel();//朗读之前先停止
    console.log("speak", timbre[1]);
  if (synth.speaking) {
    console.error("speechSynthesis.speaking");
    return;
  }
  if (timbre[1] == "") return;
  var utterThis = new SpeechSynthesisUtterance(timbre[1]);
  utterThis.onend = function (event) {
    console.log("SpeechSynthesisUtterance.onend");
	$.get("/onend",function(data,status){});//阅读完成
  };
  utterThis.onerror = function (event) {
    console.error("SpeechSynthesisUtterance.onerror");
	$.get("/onend",function(data,status){});
  };
 //utterThis.voice = synth.getVoices()[35];
    voices = synth.getVoices();//获取音色
    for (i = 0; i < voices.length; i++) {
        if (voices[i].name === timbre[0]) {
            utterThis.voice = voices[i];
        }
    }
  //utterThis.pitch = 1;
  //utterThis.rate = 1;
  synth.speak(utterThis);
  
}


//url转blob

function urlToBlob() {
  return new Promise((resolve, reject) => {
    let file_url = "readContent.txt";
    let xhr = new XMLHttpRequest();
    xhr.open("get", file_url, true);
    xhr.responseType = "blob";
    xhr.onload = function () {
      if (this.status == 200) {
        // if (callback) {
        // callback();
        console.log(this.response);
        const reader = new FileReader();
          reader.onload = function () {
              inputText = reader.result;
              console.log(reader.result,'result');
              resolve(reader.result);
        };
        reader.readAsText(this.response);
      }
    };
    xhr.send();
  });
}

function sleep(numberMillis) {
  var now = new Date();
  var exitTime = now.getTime() + numberMillis;
  while (true) {
    now = new Date();
    if (now.getTime() > exitTime) return;
  }
}

function task4() {
  return new Promise((resolve) => {
    setTimeout(() => {
      urlToBlob();
      resolve("done");
    }, 2000);
  });
}

var currText = ""; //当前文件中的内容



setInterval(() => {
  let promise = urlToBlob()
  promise.then(res => {
    console.log(res)
      if (res !== spkText) {
          var tmp = res.split("@*@");
          spkText = res;
          console.log('change');
          speak(tmp);
    }
  })
}, 1500);


