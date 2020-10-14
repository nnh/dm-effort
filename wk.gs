function setChartOptions(){
  const ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const targetCharts = ss.getCharts();
  if (targetCharts.length > 1){
    console.log('test');
    return;
  }
  var targetChart = targetCharts[0];
  const sliceColors = editGrayScaleArray();
  targetChart = targetChart.modify()
    .setOption('sliceVisibilityThreshold', .2) 
    .setOption('title', '')
    /*.setOption('slices', {0: {color: grayScale[1]}, 
                          1: {color: grayScale[2]}, 
                          2: {color: grayScale[3]}, 
                          3: {color: grayScale[4]}, 
                          4: {color: grayScale[5]}, 
                          5: {color: grayScale[6]}, 
                          6: {color: grayScale[7]}, 
                          7: {color: grayScale[8]}, 
                          8: {color: grayScale[9]}, 
                          9: {color: grayScale[8]}, 
                          10: {color: grayScale[7]}, 
                          11: {color: grayScale[6]}, 
                          12: {color: grayScale[5]}, 
                          13: {color: grayScale[4]}, 
                          14: {color: grayScale[3]}
                         })*/
    .setOption('slices', sliceColors)
    .build();
  ss.updateChart(targetChart);
}

function editGrayScaleArray(){
  const itemCount = 24;
  const grayScale = ['#FFFFFF', '#EFEFEF', '#DCDDDD', '#C9CACA', '#B5B5B6', '#9FA0A0', '#898989', '#727171', '#595757', '#3E3A39', '#231815'];
  const seqArray = [...Array(itemCount).keys()];
  const array1_9 = [...Array(9).keys()].map(x => ++x);
  const array9_1 = [...array1_9].reverse().filter((x, idx) => (idx > 0 && idx < 8));
  const arrayIdx = array1_9.concat(array9_1, array1_9, array9_1, array1_9);
  const colorArrayIdx = arrayIdx.slice(0, itemCount);
  const colorArray = colorArrayIdx.map(function(x){
    var temp = {};
    temp.color = grayScale[x];
    return temp;
  });
  const array = seqArray.map((x, idx) => [x, colorArray[idx]]);
  return Object.fromEntries(array);
}
