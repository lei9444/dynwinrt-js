use windows::core::HSTRING;
use windows_future::IAsyncOperation;

pub async fn get_async_string(
  op_string: IAsyncOperation<HSTRING>,
) -> windows::core::Result<String> {
  let s = op_string.await?;
  Ok(s.to_string())
}
