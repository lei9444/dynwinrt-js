import test from 'ava'

import {
  useDynwinrtAdd,
  getComputerName,
  getWindowsDirectory,
  ComUri,
  getUriVtable,
  callMethod,
  httpClientGetSync,
  asyncProgressHstringToPromiseString,
  initWinappsdk,
  DynWinRtValue,
  WinIIds,
  DynWinRtType,
  JsTypeRegistry,
  WinGuid,
  roInitializeMta,
  geopointGetLatitude,
  geopointGetLongitude,
} from '../dist/index.js'

test('sync function from native code', (t) => {
  t.is(useDynwinrtAdd(40, 2), 42)
  t.is(useDynwinrtAdd(5, 3), 8)
})

test('getComputerName', (t) => {
  const name = getComputerName()
  console.log('Computer name:', name)
  t.truthy(name)
  t.is(typeof name, 'string')
})

test('getWindowsDirectory', (t) => {
  const dir = getWindowsDirectory()
  console.log('Windows directory:', dir)
  t.truthy(dir)
  t.is(typeof dir, 'string')
})

test('WinRT URI and VTable', (t) => {
  const uri = ComUri.createTestUri('https://www.example.com/path?query=1#fragment')
  t.truthy(uri)

  const vtable = getUriVtable()
  t.truthy(vtable)

  const runtimeClassName = callMethod(vtable, 4, uri)
  console.log('RuntimeClassName', runtimeClassName)
  t.is(runtimeClassName, 'Windows.Foundation.Uri')

  const domain = callMethod(vtable, 13, uri)
  console.log('Domain (index 13)', domain)
  // Index 13 appears to be Path in this environment
  t.is(domain, '/path')

  const absoluteUri = callMethod(vtable, 17, uri)
  console.log('AbsoluteUri (index 17)', absoluteUri)
  t.is(absoluteUri, 'https')
})

test('WinRTType build', (t) => {
  t.truthy(DynWinRtType.i32())
  t.truthy(DynWinRtType.hstring())
  t.truthy(DynWinRtType.object())
  t.truthy(DynWinRtType.iAsyncOperation(WinIIds.iAsyncOperationPickFileResultIid()))
})

test('HttpClient Sync', async (t) => {
  const asyncOperation = httpClientGetSync('https://www.microsoft.com')
  t.truthy(asyncOperation)

  const value = await asyncProgressHstringToPromiseString(asyncOperation)
  t.truthy(value)
  t.is(typeof value, 'string')
})

// test('run picker test in Rust', async (t) => {
//   await runTestPicker()
//   t.pass()
// })

test('file open picker', async (t) => {
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
  t.pass()
})

// --- Value Type Registry Tests ---

test('TypeRegistry: create primitive types', (t) => {
  const reg = new JsTypeRegistry()
  const f32 = reg.f32Type()
  const f64 = reg.f64Type()
  const i32 = reg.i32Type()
  const u8 = reg.u8Type()

  t.is(f32.sizeOf(), 4)
  t.is(f64.sizeOf(), 8)
  t.is(i32.sizeOf(), 4)
  t.is(u8.sizeOf(), 1)

  t.is(f32.alignOf(), 4)
  t.is(f64.alignOf(), 8)
})

test('TypeRegistry: define Point struct {f32, f32}', (t) => {
  const reg = new JsTypeRegistry()
  const f32 = reg.f32Type()
  const point = reg.defineStruct([f32, f32])

  t.is(point.sizeOf(), 8)
  t.is(point.alignOf(), 4)
  t.is(point.fieldCount(), 2)
  t.is(point.fieldOffset(0), 0)
  t.is(point.fieldOffset(1), 4)
})

test('TypeRegistry: define BasicGeoposition struct {f64, f64, f64}', (t) => {
  const reg = new JsTypeRegistry()
  const f64 = reg.f64Type()
  const geo = reg.defineStruct([f64, f64, f64])

  t.is(geo.sizeOf(), 24)
  t.is(geo.alignOf(), 8)
  t.is(geo.fieldCount(), 3)
  t.is(geo.fieldOffset(0), 0)
  t.is(geo.fieldOffset(1), 8)
  t.is(geo.fieldOffset(2), 16)
})

test('ValueTypeData: set and get f32 fields', (t) => {
  const reg = new JsTypeRegistry()
  const f32 = reg.f32Type()
  const point = reg.defineStruct([f32, f32])

  const val = point.defaultValue()
  val.setF32(0, 10.5)
  val.setF32(1, 20.25)

  t.is(val.getF32(0), Math.fround(10.5))
  t.is(val.getF32(1), Math.fround(20.25))
})

test('ValueTypeData: set and get f64 fields', (t) => {
  const reg = new JsTypeRegistry()
  const f64 = reg.f64Type()
  const geo = reg.defineStruct([f64, f64, f64])

  const val = geo.defaultValue()
  val.setF64(0, 47.643)
  val.setF64(1, -122.131)
  val.setF64(2, 100.0)

  t.is(val.getF64(0), 47.643)
  t.is(val.getF64(1), -122.131)
  t.is(val.getF64(2), 100.0)
})

test('ValueTypeData: set and get i32 fields', (t) => {
  const reg = new JsTypeRegistry()
  const i32 = reg.i32Type()
  const s = reg.defineStruct([i32, i32])

  const val = s.defaultValue()
  val.setI32(0, 42)
  val.setI32(1, -7)

  t.is(val.getI32(0), 42)
  t.is(val.getI32(1), -7)
})

test('ValueTypeData: default value is zeroed', (t) => {
  const reg = new JsTypeRegistry()
  const f64 = reg.f64Type()
  const geo = reg.defineStruct([f64, f64, f64])

  const val = geo.defaultValue()
  t.is(val.getF64(0), 0.0)
  t.is(val.getF64(1), 0.0)
  t.is(val.getF64(2), 0.0)
})

test('TypeRegistry: mixed field alignment {u8, i32, u8}', (t) => {
  const reg = new JsTypeRegistry()
  const u8 = reg.u8Type()
  const i32 = reg.i32Type()
  const s = reg.defineStruct([u8, i32, u8])

  t.is(s.sizeOf(), 12)
  t.is(s.alignOf(), 4)
  t.is(s.fieldOffset(0), 0)
  t.is(s.fieldOffset(1), 4)
  t.is(s.fieldOffset(2), 8)
})

test('TypeRegistry: nested struct', (t) => {
  const reg = new JsTypeRegistry()
  const f32 = reg.f32Type()
  const f64 = reg.f64Type()
  const inner = reg.defineStruct([f32, f32])
  const outer = reg.defineStruct([inner, f64])

  t.is(outer.sizeOf(), 16)
  t.is(outer.alignOf(), 8)
  t.is(outer.fieldOffset(0), 0)
  t.is(outer.fieldOffset(1), 8)
})

test('TypeHandle: field_type returns correct type', (t) => {
  const reg = new JsTypeRegistry()
  const f32 = reg.f32Type()
  const f64 = reg.f64Type()
  const s = reg.defineStruct([f32, f64])

  t.is(s.fieldType(0).sizeOf(), 4)
  t.is(s.fieldType(1).sizeOf(), 8)
})

test('Geopoint: create via struct value type from JS', (t) => {
  roInitializeMta()

  // 1. Define BasicGeoposition { Latitude: f64, Longitude: f64, Altitude: f64 }
  const reg = new JsTypeRegistry()
  const f64 = reg.f64Type()
  const geoType = reg.defineStruct([f64, f64, f64])

  // 2. Create & fill value
  const geo = geoType.defaultValue()
  geo.setF64(0, 47.643)   // Latitude
  geo.setF64(1, -122.131)  // Longitude
  geo.setF64(2, 100.0)     // Altitude

  // 3. Get IGeopointFactory via activation factory + cast
  const factory = DynWinRtValue.activationFactory('Windows.Devices.Geolocation.Geopoint')
  const geopointFactoryIid = WinGuid.parse('DB6B8D33-76BD-4E30-8AF7-A844DC37B7A0')
  const geopointFactory = factory.cast(geopointFactoryIid)

  // 4. Call IGeopointFactory::Create(BasicGeoposition) -> Geopoint (vtable index 6)
  const geopoint = geo.callMethod(geopointFactory, 6)
  t.truthy(geopoint)

  // 5. Verify the returned object is a valid Geopoint by reading its RuntimeClassName
  // IInspectable::GetRuntimeClassName is vtable index 4
  const className = geopoint.callSingleOut0(4, DynWinRtType.hstring())
  t.is(className.toString(), 'Windows.Devices.Geolocation.Geopoint')

  // 6. Verify Position round-trips correctly (ad hoc helpers use static projection)
  const lat = geopointGetLatitude(geopoint)
  const lon = geopointGetLongitude(geopoint)
  console.log('Geopoint Latitude:', lat)
  console.log('Geopoint Longitude:', lon)
  t.true(Math.abs(lat - 47.643) < 1e-6)
  t.true(Math.abs(lon - (-122.131)) < 1e-6)
})
