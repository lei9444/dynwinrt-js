import {
    initWinappsdk,
    DynWinRtValue,
    WinIIds,
    DynWinRtType,
    OcrDemo,
    hasPackageIdentity,
} from '../dist/index.js'

async function createRecognizer() {
    const factory = DynWinRtValue.activationFactory('Microsoft.Windows.AI.Imaging.TextRecognizer').cast(WinIIds.iTextRecognizerStaticsIid());
    const readyState = factory.call0(6, DynWinRtType.i32()).toNumber();
    console.log('TextRecognizer ready state:', readyState);

    if (readyState !== 0) {
        throw new Error('TextRecognizer is not ready');
    }
    const recognizer = await factory.callSingleOut0(
        8,
        DynWinRtType.iAsyncOperation(WinIIds.iAsyncOperationTextRecognizerIid())
    ).toPromise();
    console.log('TextRecognizer created successfully');
    return recognizer;
}

function createImageBuffer(bitmap: DynWinRtValue) {
    const factory = DynWinRtValue.activationFactory('Microsoft.Graphics.Imaging.ImageBuffer').cast(WinIIds.iImageBufferStaticsIid());
    const imageBuffer = factory.callSingleOut1(
        7,
        DynWinRtType.object(),
        bitmap
    );
    console.log('ImageBuffer created from bitmap successfully');
    return imageBuffer;
}

async function main() {
    // initWinappsdk(1, 8); // no need call MddBootstrap in packaged app
    console.log(hasPackageIdentity() ? 'Has package identity' : 'No package identity');
    const image = OcrDemo.pickImage();
    const recognizer  = await createRecognizer();
    const imageBuffer = createImageBuffer(image);
    const result = recognizer.callSingleOut1(
        7,
        DynWinRtType.object(),
        imageBuffer);
    OcrDemo.printOcrLines(result);
}
main();