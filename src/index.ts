#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { encode } from "@toon-format/toon";
import { authorize } from "./auth.js";
import { gmail as googleGmail } from "@googleapis/gmail";
import { extractHeaders, extractBody, extractHtmlBody, stripHtml, buildRawMessage, extractAttachments, getAttachment } from "./gmail.js";
import { existsSync, mkdirSync, writeFileSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";

const credentialsPath = process.env.GOOGLE_OAUTH_CREDENTIALS;
if (!credentialsPath) {
  console.error("GOOGLE_OAUTH_CREDENTIALS 環境変数を設定してください");
  process.exit(1);
}
if (!existsSync(credentialsPath)) {
  console.error(`credentials.json が見つかりません: ${credentialsPath}`);
  process.exit(1);
}

const resolvedCredentialsPath: string = credentialsPath;
const resolvedTokensPath: string = process.env.GOOGLE_OAUTH_TOKENS ?? join(homedir(), ".config", "gmail-mcp", "tokens.json");

// lazy auth: ツール呼び出し時に初めて認証する
let gmailClient: ReturnType<typeof googleGmail> | null = null;

async function getGmail() {
  if (!gmailClient) {
    const auth = await authorize(resolvedCredentialsPath, resolvedTokensPath);
    gmailClient = googleGmail({ version: "v1", auth });
  }
  return gmailClient;
}

const server = new McpServer({
  name: "gmail-mcp",
  version: "0.2.0",
});

// 1. search-messages
server.registerTool(
  "search-messages",
  {
    description: "Gmail検索クエリでメール一覧を取得する。レスポンスはTOON形式で返す。",
    inputSchema: {
      query: z.string().describe("Gmail検索クエリ（例: from:user@example.com, subject:会議, is:unread など）"),
      maxResults: z.number().optional().default(20).describe("最大取得件数"),
    },
  },
  async ({ query, maxResults }) => {
    const gmail = await getGmail();
    const res = await gmail.users.messages.list({
      userId: "me",
      q: query,
      maxResults,
    });

    const messageIds = res.data.messages ?? [];
    if (messageIds.length === 0) {
      return {
        content: [{ type: "text", text: "メッセージが見つかりませんでした" }],
      };
    }

    const details = await Promise.all(
      messageIds
        .filter((msg) => msg.id)
        .map((msg) =>
          gmail.users.messages.get({
            userId: "me",
            id: msg.id!,
            format: "metadata",
            metadataHeaders: ["From", "To", "Cc", "Subject", "Date"],
          })
        )
    );

    const rows = details.map((detail) => {
      const headers = extractHeaders(detail.data.payload?.headers);
      return {
        date: headers.date,
        from: headers.from,
        to: headers.to,
        cc: headers.cc,
        subject: headers.subject,
        snippet: detail.data.snippet ?? "",
        id: detail.data.id ?? "",
        threadId: detail.data.threadId ?? "",
        labels: (detail.data.labelIds ?? []).join(", "),
      };
    });

    return {
      content: [{
        type: "text",
        text: encode({ messages: rows }),
      }],
    };
  }
);

// 2. get-messages
server.registerTool(
  "get-messages",
  {
    description: "複数のメッセージIDで本文を含むメール詳細を一括取得する。",
    inputSchema: {
      messageIds: z.array(z.string()).describe("メッセージIDの配列"),
    },
  },
  async ({ messageIds }) => {
    const gmail = await getGmail();
    const details = await Promise.all(
      messageIds.map((id) =>
        gmail.users.messages.get({
          userId: "me",
          id,
          format: "full",
        })
      )
    );

    const messages = details.map((detail) => {
      const headers = extractHeaders(detail.data.payload?.headers);
      let body = extractBody(detail.data.payload ?? undefined);
      const htmlBody = extractHtmlBody(detail.data.payload ?? undefined);
      if (!body && htmlBody) {
        body = stripHtml(htmlBody);
      }
      const attachments = extractAttachments(detail.data.payload ?? undefined);
      return {
        id: detail.data.id ?? "",
        threadId: detail.data.threadId ?? "",
        labels: (detail.data.labelIds ?? []).join(", "),
        date: headers.date,
        from: headers.from,
        to: headers.to,
        cc: headers.cc,
        subject: headers.subject,
        body,
        htmlBody,
        attachments,
      };
    });

    return {
      content: [{
        type: "text",
        text: encode({ messages }),
      }],
    };
  }
);

// 3. get-threads
server.registerTool(
  "get-threads",
  {
    description: "複数のスレッドIDでスレッド全体のメッセージを一括取得する。",
    inputSchema: {
      threadIds: z.array(z.string()).describe("スレッドIDの配列"),
    },
  },
  async ({ threadIds }) => {
    const gmail = await getGmail();
    const threads = await Promise.all(
      threadIds.map((id) =>
        gmail.users.threads.get({
          userId: "me",
          id,
          format: "full",
        })
      )
    );

    const result = threads.map((thread) => {
      const messages = (thread.data.messages ?? []).map((msg) => {
        const headers = extractHeaders(msg.payload?.headers);
        let body = extractBody(msg.payload ?? undefined);
        const htmlBody = extractHtmlBody(msg.payload ?? undefined);
        if (!body && htmlBody) {
          body = stripHtml(htmlBody);
        }
        const attachments = extractAttachments(msg.payload ?? undefined);
        return {
          id: msg.id ?? "",
          date: headers.date,
          from: headers.from,
          to: headers.to,
          cc: headers.cc,
          subject: headers.subject,
          labels: (msg.labelIds ?? []).join(", "),
          body,
          htmlBody,
          attachments,
        };
      });
      return {
        threadId: thread.data.id ?? "",
        messages,
      };
    });

    return {
      content: [{
        type: "text",
        text: encode({ threads: result }),
      }],
    };
  }
);

// 4. create-draft
server.registerTool(
  "create-draft",
  {
    description: "メールの下書きを作成する。返信の場合はthreadIdとinReplyToMessageIdを指定する。",
    inputSchema: {
      to: z.array(z.string()).describe("宛先メールアドレスの配列"),
      cc: z.array(z.string()).optional().default([]).describe("CCメールアドレスの配列"),
      subject: z.string().describe("件名"),
      body: z.string().describe("本文（プレーンテキスト）"),
      htmlBody: z.string().optional().describe("HTML本文（指定時はmultipart/alternativeで送信。返信時は引用付きHTMLを含める）"),
      threadId: z.string().optional().describe("返信先スレッドID（返信時に指定）"),
      inReplyToMessageId: z.string().optional().describe("返信先メッセージID（返信時に指定。Referencesヘッダー構築用）"),
    },
  },
  async ({ to, cc, subject, body, htmlBody, threadId, inReplyToMessageId }) => {
    const gmail = await getGmail();
    // 返信時のヘッダー構築
    let inReplyTo: string | undefined;
    let references: string | undefined;

    if (inReplyToMessageId && threadId) {
      try {
        const origMsg = await gmail.users.messages.get({
          userId: "me",
          id: inReplyToMessageId,
          format: "metadata",
          metadataHeaders: ["Message-ID", "References"],
        });
        const origHeaders = origMsg.data.payload?.headers ?? [];
        const messageIdHeader = origHeaders.find((h) => h.name?.toLowerCase() === "message-id")?.value ?? "";
        const referencesHeader = origHeaders.find((h) => h.name?.toLowerCase() === "references")?.value ?? "";

        if (messageIdHeader) {
          inReplyTo = messageIdHeader;
          references = referencesHeader ? `${referencesHeader} ${messageIdHeader}` : messageIdHeader;
        }
      } catch {
        // 元メッセージの取得に失敗しても下書き作成は続行
      }
    }

    const raw = buildRawMessage(to, cc, subject, body, threadId, inReplyTo, references, htmlBody);

    const draft = await gmail.users.drafts.create({
      userId: "me",
      requestBody: {
        message: {
          raw,
          ...(threadId && { threadId }),
        },
      },
    });

    return {
      content: [{
        type: "text",
        text: `下書きを作成しました (ID: ${draft.data.id})`,
      }],
    };
  }
);

// 5. modify-labels
server.registerTool(
  "modify-labels",
  {
    description: "メッセージのラベルを追加/削除する。アーカイブはremoveLabelIdsに\"INBOX\"を指定する。",
    inputSchema: {
      messageIds: z.array(z.string()).describe("メッセージIDの配列"),
      addLabelIds: z.array(z.string()).optional().default([]).describe("追加するラベルIDの配列"),
      removeLabelIds: z.array(z.string()).optional().default([]).describe("削除するラベルIDの配列"),
    },
  },
  async ({ messageIds, addLabelIds, removeLabelIds }) => {
    const gmail = await getGmail();
    await gmail.users.messages.batchModify({
      userId: "me",
      requestBody: {
        ids: messageIds,
        addLabelIds,
        removeLabelIds,
      },
    });

    return {
      content: [{
        type: "text",
        text: `${messageIds.length}件のメッセージのラベルを更新しました。`,
      }],
    };
  }
);

// 6. list-labels
server.registerTool(
  "list-labels",
  {
    description: "利用可能なラベル一覧を取得する。レスポンスはTOON形式で返す。",
    inputSchema: {},
  },
  async () => {
    const gmail = await getGmail();
    const res = await gmail.users.labels.list({ userId: "me" });
    const labels = (res.data.labels ?? []).map((label) => ({
      id: label.id ?? "",
      name: label.name ?? "",
      type: label.type ?? "",
    }));

    return {
      content: [{
        type: "text",
        text: encode({ labels }),
      }],
    };
  }
);

// 7. save-attachment
server.registerTool(
  "save-attachment",
  {
    description: "Download an email attachment and save it to disk.",
    inputSchema: {
      messageId: z.string().describe("The message ID containing the attachment"),
      attachmentId: z.string().describe("The attachment ID from get-messages/get-threads"),
      filename: z.string().describe("The original filename of the attachment"),
      savePath: z.string().describe("Absolute path to directory where the file will be saved"),
    },
  },
  async ({ messageId, attachmentId, filename, savePath }) => {
    const gmail = await getGmail();
    const data = await getAttachment(gmail, messageId, attachmentId);

    if (!data) {
      return {
        content: [{ type: "text", text: "Attachment data is empty." }],
      };
    }

    // Ensure save directory exists
    mkdirSync(savePath, { recursive: true });

    const filePath = join(savePath, filename);
    const buffer = Buffer.from(data, "base64url");
    writeFileSync(filePath, buffer);

    return {
      content: [{
        type: "text",
        text: `Attachment saved to ${filePath} (${buffer.length} bytes)`,
      }],
    };
  }
);

const transport = new StdioServerTransport();
await server.connect(transport);
