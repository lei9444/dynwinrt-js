import {
    DynWinRtValue,
    WinIIds,
    DynWinRtType,
    OcrDemo,
    hasPackageIdentity,
} from '../dist/index.js'

async function coreAPIDemos(obj: DynWinRtValue) {
    const factory = DynWinRtValue.activationFactory('<RuntimeClassName>')
    const primitiveValue = DynWinRtValue.i32(42)
    const casted = obj.cast("some-interface-iid");
    const asyncv = await obj.toPromise(); // async integraition
    const value = obj.callSingleOut1(3, // method index in vtable
                                     DynWinRtType.hstring(), // return type
                                     obj // argument
                                    );
    
    const typ = DynWinRtType.i32(); // primitive type
    const gtype = DynWinRtType.iAsyncOperation("some-interface-iid"); // generic type                 
}