// =============================================================================
// Cloudflare Worker: Anthropic API プロキシ（APIキー埋め込み + 合言葉認証）
// =============================================================================
// 【セットアップ】
//   1. 下の2箇所を書き換える
//   2. Cloudflare Workers にデプロイ
// =============================================================================

// ★★★ 設定（この2箇所を書き換えてください）★★★
const ANTHROPIC_API_KEY = 'ここにAPIキーを設定';
const ACCESS_TOKEN = 'ここに合言葉を設定';
// ★★★ 設定ここまで ★★★

export default {
  async fetch(request, env) {
    if (request.method === 'OPTIONS') {
      return new Response(null, {
        headers: {
          'Access-Control-Allow-Origin': '*',
          'Access-Control-Allow-Methods': 'POST, OPTIONS',
          'Access-Control-Allow-Headers': 'Content-Type, x-access-token, anthropic-version',
          'Access-Control-Max-Age': '86400',
        },
      });
    }
    if (request.method !== 'POST') {
      return new Response('Method not allowed', { status: 405 });
    }
    const token = request.headers.get('x-access-token');
    if (!token || token !== ACCESS_TOKEN) {
      return new Response(JSON.stringify({ error: 'アクセスが拒否されました。' }), {
        status: 403,
        headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' },
      });
    }
    try {
      const headers = new Headers();
      headers.set('Content-Type', 'application/json');
      headers.set('anthropic-version', request.headers.get('anthropic-version') || '2023-06-01');
      headers.set('x-api-key', ANTHROPIC_API_KEY);
      const response = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: headers,
        body: request.body,
      });
      const responseBody = await response.text();
      return new Response(responseBody, {
        status: response.status,
        headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' },
      });
    } catch (error) {
      return new Response(JSON.stringify({ error: error.message }), {
        status: 500,
        headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' },
      });
    }
  },
};
