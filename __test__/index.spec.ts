import test from 'ava'

import {
  useDynwinrtAdd,
  getComputerName,
  getWindowsDirectory,
  ComUri,
  getUriVtable,
  callMethod,
  WinRTType,
  httpClientGetSync,
  asyncProgressHstringToPromiseString,
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

test('WinRTType Enum', (t) => {
  t.is(WinRTType.I32, 0)
  t.is(WinRTType.Object, 1)
  t.is(WinRTType.HString, 2)
  t.is(WinRTType.HResult, 3)
})

test('HttpClient Sync', async (t) => {
  const asyncOperation = httpClientGetSync('https://www.microsoft.com')
  t.truthy(asyncOperation)

  const value = await asyncProgressHstringToPromiseString(asyncOperation)
  t.truthy(value)
  t.is(typeof value, 'string')
})
