importScripts('dbf-reader.js');

self.onmessage = async (e) => {
    const {file} = e.data;

    const arrayBuffer = await file.arrayBuffer();
    const data = parseDBF(arrayBuffer, 'GBK');

    self.postMessage({
        type: 'finished',
        data: data,
    });
};