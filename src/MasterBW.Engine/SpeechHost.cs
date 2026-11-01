using System.Text;
using MasterBW.Core.Util;
using Whisper.net;

namespace MasterBW.Engine;

/// <summary>Pengenalan suara offline (Whisper) untuk perintah suara bahasa Indonesia.</summary>
public sealed class SpeechHost : IDisposable
{
    private readonly string _modelPath;
    private readonly SemaphoreSlim _gate = new(1, 1);
    private WhisperFactory? _factory;

    public SpeechHost(string modelPath) { _modelPath = modelPath; }

    public bool Available => File.Exists(_modelPath);

    public async Task<string> TranscribeAsync(Stream wav, string language, CancellationToken ct)
    {
        if (!Available) throw new FileNotFoundException("Model suara tidak ditemukan: " + _modelPath + ". Unduh ggml-base.bin lewat skrip unduh model.");
        await _gate.WaitAsync(ct);
        try
        {
            _factory ??= WhisperFactory.FromPath(_modelPath);
            using var processor = _factory.CreateBuilder().WithLanguage(string.IsNullOrWhiteSpace(language) ? "id" : language).Build();
            var sb = new StringBuilder();
            await foreach (var seg in processor.ProcessAsync(wav, ct)) sb.Append(seg.Text);
            var text = sb.ToString().Trim();
            Log.Info("Transkripsi: " + text);
            return text;
        }
        finally { _gate.Release(); }
    }

    public void Dispose() => _factory?.Dispose();
}
