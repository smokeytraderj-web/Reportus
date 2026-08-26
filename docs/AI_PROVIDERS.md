# AI provider configuration

Reporticles keeps report logic independent from any model vendor. AI synthesis is disabled by default; deterministic JSON workflows continue to work without a model.

## Free local development with Ollama

1. Install Ollama from <https://ollama.com/download>.
2. Pull a model that supports structured output.
3. Set these Windows environment variables, replacing `<installed-model>` with the model name you pulled:

```powershell
setx REPORTICLES_AI_PROVIDER ollama
setx REPORTICLES_OLLAMA_MODEL "<installed-model>"
setx REPORTICLES_OLLAMA_ENDPOINT "http://127.0.0.1:11434"
```

Restart Reporticles after changing the variables. Local Ollama needs no API key and uploaded content stays on the computer. Reporticles uses Ollama's JSON-schema `format` support with non-streaming output, as documented in the [Ollama Generate API](https://docs.ollama.com/api/generate).

## Remote endpoints

A non-loopback Ollama endpoint is treated as external processing and remains blocked unless `REPORTICLES_ALLOW_EXTERNAL_AI=true` is explicitly configured. The privacy gate still runs first. Production provider selection and key management will move into the deferred admin settings area. Legacy `REPORTUS_*` variables remain supported for existing installations.
