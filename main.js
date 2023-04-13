
function fixdata(data) { //文件流转BinaryString
    var o = "",
    l = 0,
    w = 10240;
    for(; l < data.byteLength / w; ++l)
      o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w, l * w + w)));
      o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w)));
    return o;
}

// const keys  = ['F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'];
const keys  = ['F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
    'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ',
    'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM'
];
// const startNum = 3;
// const endNum = 6;
const startNum = 3;
const endNum = 203;
let accVal = 0;
let resultCol = '';
const target = -10;
let minVal = Number.MAX_VALUE;
let minValCol = '';


function fileSelectHandle () {
    let wb;  // 读取完成的数据
    let rABS = false; // 是否将文件读取为二进制字符串
    document.getElementById("uploadExcel").addEventListener("change", function(e){
        if(!e.target.files) return;
        let f = e.target.files[0];
        let reader = new FileReader();
        reader.onload = function(e) {
            let data = e.target.result;
            if(rABS) {
                wb = XLSX.read(btoa(fixdata(data)), {type: 'base64'});//手动转化
            }
            else {
                wb = XLSX.read(data, {type: 'binary'});
            }
            const sheet = wb.Sheets['吸波数据'];

            keys.forEach((key, index) => {
                let n = 0;
                for (let i = startNum; i <= endNum; i++) {
                    const val = Number(sheet[key+i].w);
                    if (val <= target) n++;
                    if (val < minVal) {
                        minVal = val;
                        minValCol = key;
                    }
                }
                if (n > accVal) {
                    accVal = n;
                    resultCol = key;
                }
            })
            const resultString = !resultCol ? `没有小于-10的值 \n 最小值：${minVal}，最小值厚度：${sheet[minValCol+2].v}` : `小于-10累计最大值：${accVal * 0.08}，最大值厚度：${sheet[resultCol+2].v}
                最小值：${minVal}，最小值厚度：${sheet[minValCol+2].v}
            `;
            console.log(resultString);
            document.getElementById('show-result').innerText = resultString;
        };
        if(rABS) {reader.readAsArrayBuffer(f);}
        else {reader.readAsBinaryString(f);}
    })
}

function eventListener() {
    fileSelectHandle();
}

function main() {
    try {
        eventListener();
    } catch (e) {
        alert('出错了，请联系 lbw！');
        console.log(e);
    }
}

main();
