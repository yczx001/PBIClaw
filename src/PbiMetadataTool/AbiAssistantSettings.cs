namespace PbiMetadataTool;

internal sealed class AbiAssistantSettings
{
    public string BaseUrl { get; set; } = "https://api.openai.com/v1";

    public string Model { get; set; } = "gpt-4.1-mini";

    public string ApiKey { get; set; } = string.Empty;

    public double Temperature { get; set; } = 0.2;

    public bool AllowModelChanges { get; set; }

    public bool IncludeHiddenObjects { get; set; }

    public string CustomSystemPrompt { get; set; } = string.Empty;

    public List<string> QuickPrompts { get; set; } =
    [
        "分析当前模型最值得优化的 3 个点，并给出可执行步骤。",
        "基于现有模型给出 5 个高价值度量值建议（含 DAX）。",
        "解释当前选中对象的业务含义与使用建议。"
    ];
}
