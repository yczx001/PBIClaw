using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;

namespace PbiMetadataTool;

internal sealed class AiChatClient
{
    private readonly HttpClient _httpClient = new();

    public async Task<string> CompleteAsync(AiChatRequest request, CancellationToken cancellationToken)
    {
        ValidateSettings(request.Settings);

        var endpoint = request.Settings.BaseUrl.TrimEnd('/') + "/chat/completions";
        var payload = BuildPayload(request);
        using var httpRequest = new HttpRequestMessage(HttpMethod.Post, endpoint)
        {
            Content = new StringContent(payload, Encoding.UTF8, "application/json")
        };
        httpRequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", request.Settings.ApiKey.Trim());

        using var response = await _httpClient.SendAsync(httpRequest, cancellationToken).ConfigureAwait(false);
        var content = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException($"AI 接口调用失败 ({(int)response.StatusCode}): {content}");
        }

        return ParseAssistantContent(content);
    }

    public async Task TestConnectionAsync(AbiAssistantSettings settings, CancellationToken cancellationToken)
    {
        ValidateSettings(settings);

        var endpoint = settings.BaseUrl.TrimEnd('/') + "/chat/completions";
        var body = new Dictionary<string, object?>
        {
            ["model"] = settings.Model.Trim(),
            ["temperature"] = 0,
            ["max_tokens"] = 1,
            ["messages"] = new[]
            {
                new Dictionary<string, object?>
                {
                    ["role"] = "user",
                    ["content"] = "ping"
                }
            }
        };

        var payload = JsonSerializer.Serialize(body);
        using var httpRequest = new HttpRequestMessage(HttpMethod.Post, endpoint)
        {
            Content = new StringContent(payload, Encoding.UTF8, "application/json")
        };
        httpRequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", settings.ApiKey.Trim());

        using var response = await _httpClient.SendAsync(httpRequest, cancellationToken).ConfigureAwait(false);
        var content = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException($"AI 接口调用失败 ({(int)response.StatusCode}): {content}");
        }
    }

    private static void ValidateSettings(AbiAssistantSettings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.BaseUrl))
        {
            throw new InvalidOperationException("Base URL 不能为空。");
        }

        if (string.IsNullOrWhiteSpace(settings.Model))
        {
            throw new InvalidOperationException("模型名称不能为空。");
        }

        if (string.IsNullOrWhiteSpace(settings.ApiKey))
        {
            throw new InvalidOperationException("API Key 不能为空。");
        }
    }

    private static string BuildPayload(AiChatRequest request)
    {
        var messages = request.Messages.Select(m => new Dictionary<string, object?>
        {
            ["role"] = m.Role,
            ["content"] = m.Content
        }).ToList();

        var body = new Dictionary<string, object?>
        {
            ["model"] = request.Settings.Model.Trim(),
            ["temperature"] = request.Settings.Temperature,
            ["messages"] = messages
        };

        return JsonSerializer.Serialize(body);
    }

    private static string ParseAssistantContent(string json)
    {
        using var doc = JsonDocument.Parse(json);
        if (!doc.RootElement.TryGetProperty("choices", out var choices) || choices.GetArrayLength() == 0)
        {
            throw new InvalidOperationException("AI 返回内容缺少 choices。");
        }

        var first = choices[0];
        if (!first.TryGetProperty("message", out var message) || !message.TryGetProperty("content", out var content))
        {
            throw new InvalidOperationException("AI 返回内容缺少 message.content。");
        }

        if (content.ValueKind == JsonValueKind.String)
        {
            return content.GetString() ?? string.Empty;
        }

        if (content.ValueKind == JsonValueKind.Array)
        {
            var textSegments = new List<string>();
            foreach (var item in content.EnumerateArray())
            {
                if (item.ValueKind == JsonValueKind.Object &&
                    item.TryGetProperty("type", out var typeElement) &&
                    string.Equals(typeElement.GetString(), "text", StringComparison.OrdinalIgnoreCase) &&
                    item.TryGetProperty("text", out var textElement))
                {
                    textSegments.Add(textElement.GetString() ?? string.Empty);
                }
            }

            if (textSegments.Count > 0)
            {
                return string.Join(Environment.NewLine, textSegments);
            }
        }

        return content.GetRawText();
    }
}

internal sealed record AiChatMessage(string Role, string Content);

internal sealed record AiChatRequest(
    AbiAssistantSettings Settings,
    IReadOnlyList<AiChatMessage> Messages);
