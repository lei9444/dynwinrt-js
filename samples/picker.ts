import {
    initWinappsdk,
    DynWinRtValue,
    WinIIds,
    DynWinRtType,
} from '../dist/index.js'

async function main() {
    initWinappsdk(1, 8)
    const factory = DynWinRtValue.activationFactory('Microsoft.Windows.Storage.Pickers.FileOpenPicker')
    const FileOpenPicker = factory.cast(WinIIds.iFileOpenPickerFactoryIid())
    const picker = FileOpenPicker.callSingleOut1(
        6,
        DynWinRtType.object(),
        DynWinRtValue.i64(0))
    const picked_file = await picker.callSingleOut0(
        13,
        DynWinRtType.iAsyncOperation(WinIIds.iAsyncOperationPickFileResultIid()))
        .toPromise()
    const path = picked_file.callSingleOut0(6, DynWinRtType.hstring());
    console.log("Selected Path", path.toString())
}
main();