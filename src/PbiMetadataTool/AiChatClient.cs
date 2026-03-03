using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;

namespace PbiMetadataTool;

internal sealed class AiChatClient
{
    private const string ProviderOpenAi = "openai";
    private const string ProviderAnthropic = "anthropic";
    private const string AnthropicApiVersion = "2023-06-01";
    private static readonly TimeSpan RequestTimeout = TimeSpan.FromSeconds(180);
    private readonly HttpClient _httpClient = new();

    public AiChatClient()
    {
        _httpClient.Timeout = RequestTimeout;
    }

    public async Task<string> CompleteAsync(AiChatRequest request, CancellationToken cancellationToken)
    {
        ValidateSettings(request.Settings);
        var provider = ResolveProvider(request.Settings);
        var payload = provider == ProviderAnthropic
            ? BuildAnthropicPayload(request)
            : BuildOpenAiPayload(request);
        var content = await PostWithEndpointFallbackAsync(
            request.Settings,
            provider,
            payload,
            cancellationToken).ConfigureAwait(false);

        return ParseAssistantContent(content, provider);
    }

    public async Task TestConnectionAsync(AbiAssistantSettings settings, CancellationToken cancellationToken)
    {
        ValidateSettings(settings);
        var provider = ResolveProvider(settings);

        Dictionary<string, object?> body;
        if (provider == ProviderAnthropic)
        {
            body = new Dictionary<string, object?>
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
        }
        else
        {
            body = new Dictionary<string, object?>
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
        }

        var payload = JsonSerializer.Serialize(body);
        _ = await PostWithEndpointFallbackAsync(
            settings,
            provider,
            payload,
            cancellationToken).ConfigureAwait(false);
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

    private static string BuildOpenAiPayload(AiChatRequest request)
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

    private static string BuildAnthropicPayload(AiChatRequest request)
    {
        var systemText = string.Join(
            Environment.NewLine + Environment.NewLine,
            request.Messages
                .Where(m => string.Equals(m.Role, "system", StringComparison.OrdinalIgnoreCase))
                .Select(m => m.Content)
                .Where(s => !string.IsNullOrWhiteSpace(s)));

        var messages = request.Messages
            .Where(m => !string.Equals(m.Role, "system", StringComparison.OrdinalIgnoreCase))
            .Select(m => new Dictionary<string, object?>
            {
                ["role"] = string.Equals(m.Role, "assistant", StringComparison.OrdinalIgnoreCase) ? "assistant" : "user",
                ["content"] = m.Content
            })
            .ToList();

        if (messages.Count == 0)
        {
            messages.Add(new Dictionary<string, object?>
            {
                ["role"] = "user",
                ["content"] = "ping"
            });
        }

        var body = new Dictionary<string, object?>
        {
            ["model"] = request.Settings.Model.Trim(),
            ["max_tokens"] = 4096,
            ["temperature"] = request.Settings.Temperature,
            ["messages"] = messages
        };

        if (!string.IsNullOrWhiteSpace(systemText))
        {
            body["system"] = systemText;
        }

        return JsonSerializer.Serialize(body);
    }

    private static string ParseAssistantContent(string json, string provider)
    {
        if (provider == ProviderAnthropic)
        {
            return ParseAnthropicContent(json);
        }

        return ParseOpenAiContent(json);
    }

    private static string ParseOpenAiContent(string json)
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

    private static string ParseAnthropicContent(string json)
    {
        using var doc = JsonDocument.Parse(json);
        if (!doc.RootElement.TryGetProperty("content", out var contentNode))
        {
            throw new InvalidOperationException("AI 返回内容缺少 content。");
        }

        if (contentNode.ValueKind == JsonValueKind.String)
        {
            return contentNode.GetString() ?? string.Empty;
        }

        if (contentNode.ValueKind == JsonValueKind.Array)
        {
            var textSegments = new List<string>();
            foreach (var item in contentNode.EnumerateArray())
            {
                if (item.ValueKind != JsonValueKind.Object)
                {
                    continue;
                }

                var type = item.TryGetProperty("type", out var typeNode) ? typeNode.GetString() : string.Empty;
                if (!string.Equals(type, "text", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (item.TryGetProperty("text", out var textNode))
                {
                    textSegments.Add(textNode.GetString() ?? string.Empty);
                }
            }

            if (textSegments.Count > 0)
            {
                return string.Join(Environment.NewLine, textSegments);
            }
        }

        return contentNode.GetRawText();
    }

    private static void ApplyAuthHeaders(HttpRequestMessage request, AbiAssistantSettings settings, string provider)
    {
        if (provider == ProviderAnthropic)
        {
            request.Headers.Add("x-api-key", settings.ApiKey.Trim());
            request.Headers.Add("anthropic-version", AnthropicApiVersion);
            // Some Claude relay services expect Bearer auth even for Anthropic-compatible routes.
            if (!IsAnthropicHost(settings.BaseUrl))
            {
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", settings.ApiKey.Trim());
            }
            return;
        }

        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", settings.ApiKey.Trim());
    }

    private async Task<string> PostWithEndpointFallbackAsync(
        AbiAssistantSettings settings,
        string provider,
        string payload,
        CancellationToken cancellationToken)
    {
        var endpoints = BuildEndpointCandidates(settings, provider);
        var failures = new List<string>();

        foreach (var endpoint in endpoints)
        {
            using var request = new HttpRequestMessage(HttpMethod.Post, endpoint)
            {
                Content = new StringContent(payload, Encoding.UTF8, "application/json")
            };
            ApplyAuthHeaders(request, settings, provider);

            using var response = await _httpClient.SendAsync(request, cancellationToken).ConfigureAwait(false);
            var content = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
            if (response.IsSuccessStatusCode)
            {
                return content;
            }

            failures.Add($"{endpoint} -> {(int)response.StatusCode}: {content}");
        }

        throw new InvalidOperationException(
            "AI 接口调用失败，已尝试以下地址：" + Environment.NewLine + string.Join(Environment.NewLine, failures));
    }

    private static IReadOnlyList<string> BuildEndpointCandidates(AbiAssistantSettings settings, string provider)
    {
        var baseUrl = settings.BaseUrl.Trim().TrimEnd('/');
        var endpoints = new List<string>();

        if (provider == ProviderAnthropic)
        {
            if (baseUrl.EndsWith("/v1/messages", StringComparison.OrdinalIgnoreCase) ||
                baseUrl.EndsWith("/messages", StringComparison.OrdinalIgnoreCase))
            {
                endpoints.Add(baseUrl);
                return endpoints;
            }

            if (baseUrl.EndsWith("/v1", StringComparison.OrdinalIgnoreCase))
            {
                endpoints.Add(baseUrl + "/messages");
                return endpoints;
            }

            // Prefer /v1/messages for relay gateways, then fallback to /messages.
            endpoints.Add(baseUrl + "/v1/messages");
            endpoints.Add(baseUrl + "/messages");
            return endpoints;
        }

        if (baseUrl.EndsWith("/chat/completions", StringComparison.OrdinalIgnoreCase))
        {
            endpoints.Add(baseUrl);
            return endpoints;
        }

        if (baseUrl.EndsWith("/v1", StringComparison.OrdinalIgnoreCase))
        {
            endpoints.Add(baseUrl + "/chat/completions");
            return endpoints;
        }

        endpoints.Add(baseUrl + "/chat/completions");
        endpoints.Add(baseUrl + "/v1/chat/completions");
        return endpoints;
    }

    private static string ResolveProvider(AbiAssistantSettings settings)
    {
        if (string.Equals(settings.Provider, ProviderAnthropic, StringComparison.OrdinalIgnoreCase))
        {
            return ProviderAnthropic;
        }

        var baseUrl = settings.BaseUrl ?? string.Empty;
        if (baseUrl.Contains("anthropic.com", StringComparison.OrdinalIgnoreCase))
        {
            return ProviderAnthropic;
        }

        return ProviderOpenAi;
    }

    private static bool IsAnthropicHost(string url)
    {
        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri))
        {
            return false;
        }

        return uri.Host.Contains("anthropic.com", StringComparison.OrdinalIgnoreCase);
    }
}

internal sealed record AiChatMessage(string Role, string Content);

internal sealed record AiChatRequest(
    AbiAssistantSettings Settings,
    IReadOnlyList<AiChatMessage> Messages);
