function exportPages(urlString, title, page=1, totalString='') {
  const fullUrlString = `${urlString}${padNumber(page, 3)}.html`;
  const xhr = new XMLHttpRequest();

  xhr.open("GET", fullUrlString, false);
  xhr.send();
  const result = xhr.responseText;
  console.log(result);
  if (xhr.status === 200 && result) {
    const div = document.createElement('div');
    div.innerHTML = result.replace('Download now', '');
    totalString += div.textContent;
    setTimeout(() => exportPages(urlString, title, ++page, totalString), 1000);
  } else {
    console.log(page);
    console.log(totalString);
    downloadString(totalString, title);
  }
}

function padNumber(number, length) {
  const array = [];
  array[length - number.toString().length] = number;
  return array.join('0');
}

function downloadString(csvString, title) {
  const csvData = new Blob([csvString]);
  const link = document.createElement('a');
  link.href =  URL.createObjectURL(csvData);
  link.target = '_blank';
  link.download = `${title}.txt`;
  link.click();
}
