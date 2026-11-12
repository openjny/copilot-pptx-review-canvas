# インストールと更新

[English](../install.md)

Copilot CLI のコマンドは [README](../../README.ja.md#インストール) にあります。このページは、それ以外の内容です。

## VS Code

このリポジトリをマーケットプレイスとして追加し、そこからプラグインをインストールします。

1. コマンドパレットで **Chat: Open Customizations** を実行し、**Browse** を選びます。
2. ソースのメニューから **Configure Marketplaces**、**Add Marketplace** の順に選び、`openjny/copilot-pptx-review-canvas` を入力します。
3. **Browse** に戻り、ソースのメニューでこのマーケットプレイスを選んで、プラグインの **Install** を押します。

## 更新

```powershell
copilot plugin update copilot-pptx-review-canvas@copilot-pptx-review-canvas
```

## 知っておくこと

- キャンバスの ID は `pptx-review-canvas` です。同じプラグインを複数の方法でインストールすると、キャンバスが二重に登録されます。インストールは 1 つの方法だけにしてください。
- 拡張機能は、メッセージを送るときに動くフックを 1 つ登録します。エージェントがまだ読んでいない書き込みがあるときに、そのことを伝える短い非表示のメモを足すためのものです。フックは特権のある機能なので、初めて使うときに Copilot のクライアントから拡張機能の承認を求められます。フックの実装は [`extension.mjs`](../../com.github.copilot/extensions/pptx-review-canvas/extension.mjs) にあります。「レビューを提出」も、同じセッションの API でエージェントにメッセージを送ります。
