using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using MasterBW.Core.Util;
using Newtonsoft.Json;

namespace MasterBW.Core.Ai
{
    public interface ILlmClient
    {
        Task<GenerateResponse> GenerateAsync(GenerateRequest request, CancellationToken ct, Action<StreamChunk> onChunk = null);
        Task<EngineStatus> GetStatusAsync(CancellationToken ct);
    }

    /// <summary>
    /// Klien HTTP ke MasterBW.Engine yang berjalan di 127.0.0.1.
    /// Satu engine dipakai bersama oleh Word, Excel, dan fungsi sel, sehingga model hanya dimuat sekali di RAM.
    /// </summary>
    public sealed class EngineClient : ILlmClient, IDisposable
    {
        private readonly HttpClient _http;
        private EngineEndpoint _endpoint;

        public EngineClient()
        {
            _http = new HttpClient { Timeout = Timeout.InfiniteTimeSpan };
        }

        public EngineEndpoint Endpoint
        {
            get
            {
                if (_endpoint == null) Refresh();
                return _endpoint;
            }
        }

        /// <summary>Baca ulang engine.json (port dan token bisa berubah setiap engine dinyalakan).</summary>
        public void Refresh()
        {
            _endpoint = AppPaths.ReadJson<EngineEndpoint>(AppPaths.EngineEndpointFile)
                        ?? new EngineEndpoint { Port = MasterBWSettings.Load().EnginePort, Token = "" };
        }

        private string BaseUrl => "http://127.0.0.1:" + Endpoint.Port;

        private HttpRequestMessage NewRequest(HttpMethod method, string path)
        {
            var req = new HttpRequestMessage(method, BaseUrl + path);
            if (!string.IsNullOrEmpty(Endpoint.Token)) req.Headers.Add("X-MasterBW-Token", Endpoint.Token);
            return req;
        }

        public async Task<EngineStatus> GetStatusAsync(CancellationToken ct)
        {
            try
            {
                Refresh();
                using (var cts = CancellationTokenSource.CreateLinkedTokenSource(ct))
                {
                    cts.CancelAfter(TimeSpan.FromSeconds(3));
                    using (var req = NewRequest(HttpMethod.Get, "/health"))
                    using (var resp = await _http.SendAsync(req, cts.Token).ConfigureAwait(false))
                    {
                        var body = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
                        return JsonConvert.DeserializeObject<EngineStatus>(body) ?? new EngineStatus { Ok = false };
                    }
                }
            }
            catch (Exception ex)
            {
                return new EngineStatus { Ok = false, Error = ex.Message };
            }
        }

        public async Task<GenerateResponse> GenerateAsync(GenerateRequest request, CancellationToken ct, Action<StreamChunk> onChunk = null)
        {
            var json = JsonConvert.SerializeObject(request);
            using (var req = NewRequest(HttpMethod.Post, "/v1/generate?stream=" + (onChunk != null ? "1" : "0")))
            {
                req.Content = new StringContent(json, Encoding.UTF8, "application/json");
                using (var resp = await _http.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct).ConfigureAwait(false))
                {
                    if (!resp.IsSuccessStatusCode)
                    {
                        var err = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
                        throw new InvalidOperationException("Engine AI menolak permintaan (" + (int)resp.StatusCode + "): " + err);
                    }

                    if (onChunk == null)
                    {
                        var body = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
                        var result = JsonConvert.DeserializeObject<GenerateResponse>(body);
                        if (!string.IsNullOrEmpty(result?.Error)) throw new InvalidOperationException(result.Error);
                        return result;
                    }

                    using (var stream = await resp.Content.ReadAsStreamAsync().ConfigureAwait(false))
                    using (var reader = new StreamReader(stream, Encoding.UTF8))
                    {
                        GenerateResponse final = null;
                        using (ct.Register(() => { try { stream.Dispose(); } catch { } }))
                        {
                            string line;
                            while ((line = await reader.ReadLineAsync().ConfigureAwait(false)) != null)
                            {
                                ct.ThrowIfCancellationRequested();
                                if (!line.StartsWith("data:", StringComparison.Ordinal)) continue;
                                var payload = line.Substring(5).Trim();
                                if (payload.Length == 0) continue;
                                var chunk = JsonConvert.DeserializeObject<StreamChunk>(payload);
                                if (chunk == null) continue;
                                if (chunk.Kind == "done")
                                {
                                    final = chunk.Result;
                                    break;
                                }
                                onChunk(chunk);
                            }
                        }
                        ct.ThrowIfCancellationRequested();
                        if (final == null) throw new InvalidOperationException("Koneksi ke engine AI terputus.");
                        if (!string.IsNullOrEmpty(final.Error)) throw new InvalidOperationException(final.Error);
                        return final;
                    }
                }
            }
        }

        /// <summary>Ubah rekaman suara (WAV 16 kHz mono) menjadi teks, sepenuhnya offline (Whisper).</summary>
        public async Task<string> TranscribeAsync(string wavPath, CancellationToken ct)
        {
            using (var req = NewRequest(HttpMethod.Post, "/v1/transcribe?lang=id"))
            {
                var bytes = File.ReadAllBytes(wavPath);
                req.Content = new ByteArrayContent(bytes);
                req.Content.Headers.ContentType = new MediaTypeHeaderValue("audio/wav");
                using (var resp = await _http.SendAsync(req, ct).ConfigureAwait(false))
                {
                    var body = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
                    var result = JsonConvert.DeserializeObject<TranscribeResponse>(body);
                    if (result == null) throw new InvalidOperationException("Jawaban engine tidak valid.");
                    if (!string.IsNullOrEmpty(result.Error)) throw new InvalidOperationException(result.Error);
                    return result.Text?.Trim() ?? "";
                }
            }
        }

        /// <summary>Minta engine melepas model dari RAM.</summary>
        public async Task UnloadAsync(CancellationToken ct)
        {
            using (var req = NewRequest(HttpMethod.Post, "/v1/unload"))
            using (await _http.SendAsync(req, ct).ConfigureAwait(false)) { }
        }

        public void Dispose() => _http.Dispose();
    }
}
