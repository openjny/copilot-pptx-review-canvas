# Install and update

[日本語](./ja/install.md)

The Copilot CLI commands are in the [README](../README.md#install). This page covers the rest.

## VS Code

Add this repository as a marketplace and install the plugin from it:

1. Run **Chat: Open Customizations** and select **Browse**.
2. Open the source menu, choose **Configure Marketplaces**, then **Add Marketplace**, and enter `openjny/copilot-pptx-review-canvas`.
3. Back in **Browse**, pick that marketplace from the source menu and select **Install** on the plugin.

## Update

```powershell
copilot plugin update copilot-pptx-review-canvas@copilot-pptx-review-canvas
```

## Good to know

- The canvas id is `pptx-review-canvas`. Install the plugin from one source only: the same plugin installed from two sources registers the canvas twice.
- The extension registers one hook that runs when you send a message: if you have marks that the agent has not read yet, it adds a short hidden note saying so. Because hooks are a privileged capability, your Copilot client asks you to approve the extension the first time. The hook is in [`extension.mjs`](../com.github.copilot/extensions/pptx-review-canvas/extension.mjs). **Submit review** sends a message to the agent through the same session API.
