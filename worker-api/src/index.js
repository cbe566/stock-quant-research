/**
 * Jamie 統一 API Gateway — Cloudflare Worker + D1
 *
 * 股票篩選端點：
 *   GET  /api/latest          — 最新一天的篩選結果
 *   GET  /api/date/:date      — 指定日期的篩選結果
 *   GET  /api/stock/:ticker   — 單隻股票歷史評分
 *   GET  /api/top/:market     — 指定市場 TOP 10 買入/賣出
 *   GET  /api/summary         — 市場概覽（最近7天）
 *   GET  /api/history/:ticker — 單隻股票得分走勢（最近30天）
 *   POST /api/upload          — GitHub Actions 上傳篩選結果（需 API Key）
 *
 * 宏觀數據端點：
 *   GET  /api/macro/latest     — 最新宏觀數據
 *   GET  /api/macro/date/:date — 指定日期宏觀數據
 *   GET  /api/macro/indicators — 指標時序（最近30天）
 *   POST /api/macro/upload     — 上傳宏觀數據（需 API Key）
 *
 * 系統端點：
 *   POST /api/email-log        — 記錄 Email 發送狀態
 */

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const path = url.pathname;

    // CORS 標頭
    const corsHeaders = {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization',
    };

    if (request.method === 'OPTIONS') {
      return new Response(null, { headers: corsHeaders });
    }

    try {
      let result;

      // 路由
      if (path === '/api/upload' && request.method === 'POST') {
        result = await handleUpload(request, env);
      } else if (path === '/api/latest') {
        result = await handleLatest(env);
      } else if (path.startsWith('/api/date/')) {
        const date = path.split('/api/date/')[1];
        result = await handleByDate(env, date);
      } else if (path.startsWith('/api/stock/')) {
        const ticker = decodeURIComponent(path.split('/api/stock/')[1]);
        result = await handleStock(env, ticker);
      } else if (path.startsWith('/api/top/')) {
        const market = decodeURIComponent(path.split('/api/top/')[1]);
        result = await handleTop(env, market);
      } else if (path === '/api/summary') {
        result = await handleSummary(env);
      } else if (path.startsWith('/api/history/')) {
        const ticker = decodeURIComponent(path.split('/api/history/')[1]);
        result = await handleHistory(env, ticker);
      // === 宏觀數據端點 ===
      } else if (path === '/api/macro/upload' && request.method === 'POST') {
        result = await handleMacroUpload(request, env);
      } else if (path === '/api/macro/latest') {
        result = await handleMacroLatest(env);
      } else if (path.startsWith('/api/macro/date/')) {
        const date = path.split('/api/macro/date/')[1];
        result = await handleMacroByDate(env, date);
      } else if (path === '/api/macro/indicators') {
        result = await handleMacroIndicators(env);
      // === 系統端點 ===
      } else if (path === '/api/email-log' && request.method === 'POST') {
        result = await handleEmailLog(request, env);
      } else if (path === '/' || path === '') {
        result = {
          service: 'Jamie 統一 API Gateway',
          version: '2.0',
          endpoints: {
            stocks: [
              'GET /api/latest',
              'GET /api/date/:date',
              'GET /api/stock/:ticker',
              'GET /api/top/:market',
              'GET /api/summary',
              'GET /api/history/:ticker',
              'POST /api/upload',
            ],
            macro: [
              'GET /api/macro/latest',
              'GET /api/macro/date/:date',
              'GET /api/macro/indicators',
              'POST /api/macro/upload',
            ],
            system: [
              'POST /api/email-log',
            ],
          },
        };
      } else {
        return new Response(JSON.stringify({ error: '找不到端點' }), {
          status: 404,
          headers: { ...corsHeaders, 'Content-Type': 'application/json' },
        });
      }

      return new Response(JSON.stringify(result, null, 2), {
        headers: { ...corsHeaders, 'Content-Type': 'application/json; charset=utf-8' },
      });
    } catch (err) {
      return new Response(JSON.stringify({ error: err.message }), {
        status: 500,
        headers: { ...corsHeaders, 'Content-Type': 'application/json' },
      });
    }
  },
};

// ==================== 上傳端點（GitHub Actions 呼叫） ====================

async function handleUpload(request, env) {
  // 驗證 API Key
  const authHeader = request.headers.get('Authorization');
  const apiKey = authHeader?.replace('Bearer ', '');
  if (apiKey !== env.API_KEY) {
    throw new Error('未授權：API Key 不正確');
  }

  const data = await request.json();
  const { date, markets } = data;

  if (!date || !markets) {
    throw new Error('缺少必要欄位：date, markets');
  }

  let totalInserted = 0;

  for (const [marketName, results] of Object.entries(markets)) {
    if (!results || results.length === 0) continue;

    // 使用 D1 batch API 批次插入（每條 SQL 1 筆，避免變數上限）
    const stmt = env.DB.prepare(`
      INSERT OR REPLACE INTO screening_results
      (date, market, ticker, name, sector, current_price,
       total_score, buy_score, sell_score,
       quality_score, value_score, momentum_score,
       tech_signal, zscore, f_score, signals)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `);

    // D1 batch 最多 500 條語句，分批處理
    const batchSize = 200;
    for (let i = 0; i < results.length; i += batchSize) {
      const batch = results.slice(i, i + batchSize);
      const stmts = batch.map(r => stmt.bind(
        date,
        marketName,
        r.ticker || '',
        r.name || '',
        r.sector || '',
        r.current_price || null,
        r.total_score || 0,
        r.buy_score || 0,
        r.sell_score || 0,
        r.quality || null,
        r.value || null,
        r.momentum || null,
        r.tech_signal || null,
        r.zscore || null,
        r.f_score || null,
        r.signals ? r.signals.join('、') : '',
      ));

      await env.DB.batch(stmts);
      totalInserted += batch.length;
    }

    // 插入市場概覽
    const totalStocks = results.length;
    const bullish = results.filter(r => (r.total_score || 0) >= 3).length;
    const bearish = results.filter(r => (r.total_score || 0) <= -3).length;
    const neutral = totalStocks - bullish - bearish;
    const avgScore = results.reduce((s, r) => s + (r.total_score || 0), 0) / totalStocks;

    const sorted = [...results].sort((a, b) => (b.total_score || 0) - (a.total_score || 0));
    const topBuy = sorted.slice(0, 5).map(r => r.ticker).join(', ');
    const topSell = sorted.slice(-5).reverse().map(r => r.ticker).join(', ');

    await env.DB.prepare(`
      INSERT OR REPLACE INTO market_summary
      (date, market, total_stocks, bullish_count, bearish_count, neutral_count, avg_score, top_buy, top_sell)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).bind(date, marketName, totalStocks, bullish, bearish, neutral, avgScore, topBuy, topSell).run();
  }

  return { success: true, date, inserted: totalInserted };
}

// ==================== 查詢端點 ====================

async function handleLatest(env) {
  const dateRow = await env.DB.prepare(
    'SELECT DISTINCT date FROM screening_results ORDER BY date DESC LIMIT 1'
  ).first();

  if (!dateRow) return { error: '尚無數據' };

  const date = dateRow.date;
  return handleByDate(env, date);
}

async function handleByDate(env, date) {
  const results = await env.DB.prepare(
    'SELECT * FROM screening_results WHERE date = ? ORDER BY market, total_score DESC'
  ).bind(date).all();

  const summary = await env.DB.prepare(
    'SELECT * FROM market_summary WHERE date = ?'
  ).bind(date).all();

  return {
    date,
    total_stocks: results.results.length,
    summary: summary.results,
    markets: groupByMarket(results.results),
  };
}

async function handleStock(env, ticker) {
  const results = await env.DB.prepare(
    'SELECT * FROM screening_results WHERE ticker = ? ORDER BY date DESC LIMIT 30'
  ).bind(ticker).all();

  return {
    ticker,
    history: results.results,
  };
}

async function handleTop(env, market) {
  const dateRow = await env.DB.prepare(
    'SELECT DISTINCT date FROM screening_results WHERE market = ? ORDER BY date DESC LIMIT 1'
  ).bind(market).first();

  if (!dateRow) return { error: '找不到該市場數據' };

  const date = dateRow.date;

  const buy = await env.DB.prepare(
    'SELECT * FROM screening_results WHERE date = ? AND market = ? ORDER BY total_score DESC LIMIT 10'
  ).bind(date, market).all();

  const sell = await env.DB.prepare(
    'SELECT * FROM screening_results WHERE date = ? AND market = ? ORDER BY total_score ASC LIMIT 10'
  ).bind(date, market).all();

  return { date, market, buy: buy.results, sell: sell.results };
}

async function handleSummary(env) {
  const results = await env.DB.prepare(
    'SELECT * FROM market_summary ORDER BY date DESC, market LIMIT 28'
  ).all();

  // 按日期分組
  const grouped = {};
  for (const row of results.results) {
    if (!grouped[row.date]) grouped[row.date] = [];
    grouped[row.date].push(row);
  }

  return { days: grouped };
}

async function handleHistory(env, ticker) {
  const results = await env.DB.prepare(
    'SELECT date, total_score, tech_signal, zscore, current_price FROM screening_results WHERE ticker = ? ORDER BY date DESC LIMIT 30'
  ).bind(ticker).all();

  return { ticker, trend: results.results };
}

// ==================== 宏觀數據端點 ====================

async function handleMacroUpload(request, env) {
  const authHeader = request.headers.get('Authorization');
  const apiKey = authHeader?.replace('Bearer ', '');
  if (apiKey !== env.API_KEY) {
    throw new Error('未授權：API Key 不正確');
  }

  const data = await request.json();
  const { date, market_data, indicators, hot_stocks, news } = data;

  if (!date) throw new Error('缺少 date 欄位');

  let inserted = 0;

  // 寫入市場數據
  if (market_data) {
    for (const [category, items] of Object.entries(market_data)) {
      if (!items || typeof items !== 'object') continue;
      for (const [symbol, info] of Object.entries(items)) {
        await env.DB.prepare(`
          INSERT OR REPLACE INTO macro_market_data (date, category, symbol, name, price, change_pct, volume, ytd_pct, extra_json)
          VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        `).bind(
          date, category, symbol,
          info.name || symbol,
          info.current || info.price || null,
          info.change_pct || null,
          info.volume || null,
          info.ytd_pct || null,
          JSON.stringify(info),
        ).run();
        inserted++;
      }
    }
  }

  // 寫入指標
  if (indicators) {
    for (const [name, info] of Object.entries(indicators)) {
      await env.DB.prepare(`
        INSERT OR REPLACE INTO macro_indicators (date, indicator_name, value, rating, extra_json)
        VALUES (?, ?, ?, ?, ?)
      `).bind(
        date, name,
        info.value || info.score || null,
        info.rating || null,
        JSON.stringify(info),
      ).run();
      inserted++;
    }
  }

  // 寫入熱門股票
  if (hot_stocks) {
    for (const [market, directions] of Object.entries(hot_stocks)) {
      for (const [direction, stocks] of Object.entries(directions)) {
        if (!Array.isArray(stocks)) continue;
        for (const s of stocks.slice(0, 10)) {
          await env.DB.prepare(`
            INSERT OR REPLACE INTO macro_hot_stocks (date, market, direction, symbol, name, current_price, change_pct, volume_ratio, in_news)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
          `).bind(
            date, market, direction,
            s.symbol || s.ticker || '',
            s.name || '',
            s.current || s.current_price || null,
            s.change_pct || null,
            s.volume_ratio || null,
            s.in_news ? 1 : 0,
          ).run();
          inserted++;
        }
      }
    }
  }

  // 寫入新聞（只存前 50 條）
  if (news && Array.isArray(news.articles)) {
    const articles = news.articles.slice(0, 50);
    for (const a of articles) {
      await env.DB.prepare(`
        INSERT INTO macro_news (date, source, tier, title, summary, url, category, sentiment, tickers_json)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
      `).bind(
        date,
        a.source || '',
        a.tier || 3,
        a.title || '',
        a.summary || '',
        a.url || '',
        a.category || '',
        a.sentiment || null,
        a.tickers ? JSON.stringify(a.tickers) : null,
      ).run();
      inserted++;
    }
  }

  return { success: true, date, inserted };
}

async function handleMacroLatest(env) {
  const dateRow = await env.DB.prepare(
    'SELECT DISTINCT date FROM macro_market_data ORDER BY date DESC LIMIT 1'
  ).first();

  if (!dateRow) return { error: '尚無宏觀數據' };
  return handleMacroByDate(env, dateRow.date);
}

async function handleMacroByDate(env, date) {
  const market = await env.DB.prepare(
    'SELECT * FROM macro_market_data WHERE date = ? ORDER BY category, symbol'
  ).bind(date).all();

  const indicators = await env.DB.prepare(
    'SELECT * FROM macro_indicators WHERE date = ?'
  ).bind(date).all();

  const hot = await env.DB.prepare(
    'SELECT * FROM macro_hot_stocks WHERE date = ? ORDER BY market, direction, change_pct DESC'
  ).bind(date).all();

  const news = await env.DB.prepare(
    'SELECT * FROM macro_news WHERE date = ? ORDER BY tier, id LIMIT 30'
  ).bind(date).all();

  return {
    date,
    market_data: groupByCategory(market.results),
    indicators: indicators.results,
    hot_stocks: groupByMarketDirection(hot.results),
    news: news.results,
  };
}

async function handleMacroIndicators(env) {
  const results = await env.DB.prepare(
    'SELECT * FROM macro_indicators ORDER BY date DESC, indicator_name LIMIT 300'
  ).all();

  const grouped = {};
  for (const row of results.results) {
    if (!grouped[row.indicator_name]) grouped[row.indicator_name] = [];
    grouped[row.indicator_name].push({ date: row.date, value: row.value, rating: row.rating });
  }

  return { indicators: grouped };
}

async function handleEmailLog(request, env) {
  const authHeader = request.headers.get('Authorization');
  const apiKey = authHeader?.replace('Bearer ', '');
  if (apiKey !== env.API_KEY) {
    throw new Error('未授權');
  }

  const data = await request.json();
  await env.DB.prepare(`
    INSERT INTO email_log (date, project, recipients_count, status, error_message)
    VALUES (?, ?, ?, ?, ?)
  `).bind(
    data.date || new Date().toISOString().slice(0, 10),
    data.project || 'unknown',
    data.recipients_count || 0,
    data.status || 'unknown',
    data.error_message || null,
  ).run();

  return { success: true };
}

function groupByCategory(rows) {
  const groups = {};
  for (const row of rows) {
    if (!groups[row.category]) groups[row.category] = [];
    groups[row.category].push(row);
  }
  return groups;
}

function groupByMarketDirection(rows) {
  const groups = {};
  for (const row of rows) {
    if (!groups[row.market]) groups[row.market] = {};
    if (!groups[row.market][row.direction]) groups[row.market][row.direction] = [];
    groups[row.market][row.direction].push(row);
  }
  return groups;
}

// ==================== 工具函數 ====================

function groupByMarket(rows) {
  const groups = {};
  for (const row of rows) {
    if (!groups[row.market]) groups[row.market] = [];
    groups[row.market].push(row);
  }
  return groups;
}
