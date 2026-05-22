export default function Page() {
  return (
    <>
      <div className="auth-overlay" id="authOverlay" role="dialog" aria-modal="true" aria-labelledby="authTitle">
        <form className="auth-card" id="authForm">
          <div>
            <h1 id="authTitle">Excel 自动生成工具</h1>
            <p>请输入访问口令后继续使用。</p>
          </div>
          <label>访问口令
            <input id="authPassword" type="password" autoComplete="current-password" placeholder="请输入访问口令" />
          </label>
          <div className="auth-error" id="authError" aria-live="polite"></div>
          <button className="primary-action" id="authSubmitBtn" type="submit">进入</button>
        </form>
      </div>
      <main className="app-shell">
      <header className="topbar">
        <div>
          <h1>Excel 自动生成工具</h1>
          <p className="meta"><span id="fieldCount">0 个表头</span><span id="rowCount">0 行数据</span></p>
        </div>
        <div className="topbar-actions">
          <button className="secondary-action" id="llmConfigBtn" type="button" aria-expanded="false" aria-controls="llmConfigPanel">大模型配置</button>
          <button className="primary-action" id="generateBtn" type="button">生成 Excel</button>
        </div>
      </header>

      <section className="llm-config-panel panel" id="llmConfigPanel" hidden aria-label="大模型配置">
        <form id="llmConfigForm" className="llm-config-form">
          <div>
            <h2>大模型配置</h2>
            <p>访问口令保存在本机浏览器，模型 API Key 保存到服务器本地配置文件，所有用户共用这一套模型配置。</p>
          </div>
          <div className="llm-config-grid">
            <label>Base URL
              <input id="llmBaseUrl" type="url" placeholder="https://ark.cn-beijing.volces.com/api/v3" required />
            </label>
            <label>模型名称
              <input id="llmModel" type="text" placeholder="请输入模型名称" required />
            </label>
            <label>API Key
              <input id="llmApiKey" type="password" placeholder="请输入 API Key" autoComplete="off" />
            </label>
            <label>超时毫秒
              <input id="llmTimeoutMs" type="number" min="5000" max="180000" step="1000" placeholder="60000" />
            </label>
          </div>
          <div className="llm-config-actions">
            <span className="llm-config-status" id="llmConfigStatus"></span>
            <button className="secondary-action" id="llmConfigCancelBtn" type="button">关闭</button>
            <button className="primary-action" type="submit">保存配置</button>
          </div>
        </form>
      </section>

      <section className="workspace" aria-label="Excel 表格生成器">
        <section className="panel panel-compact config-panel" id="configPanel" aria-labelledby="configTitle">
          <div className="panel-heading">
            <h2 id="configTitle">表头配置</h2>
            <div className="heading-actions">
              <button className="config-toggle" id="toggleConfigBtn" type="button" aria-expanded="false" aria-controls="fieldList">
                <span className="config-toggle-text">展开</span>
                <span className="config-toggle-icon" aria-hidden="true">⌄</span>
              </button>
              <button className="icon-button" id="addFieldBtn" type="button" aria-label="添加表头" title="添加表头">+</button>
            </div>
          </div>
          <div className="field-list" id="fieldList">
            <div className="empty-state">页面正在加载，如果手机上长时间停留在这里，请换用系统浏览器并刷新。</div>
          </div>
        </section>

        <section className="panel panel-main" aria-labelledby="dataTitle">
          <div className="panel-heading">
            <h2 id="dataTitle">填写数据</h2>
            <div className="toolbar">
              <button className="secondary-action" id="clearRowsBtn" type="button">清空行</button>
              <button className="secondary-action" id="addCustomRowTopBtn" type="button">自由行</button>
              <button className="icon-button" id="addRowTopBtn" type="button" aria-label="添加行" title="添加行">+</button>
            </div>
          </div>
          <div className="draft-panel" id="draftPanel">
            <div className="draft-unavailable">正在读取本机草稿...</div>
          </div>
          <div className="document-name-panel">
            <label>名称
              <input id="documentName" type="text" placeholder="请输入名称" />
            </label>
            <label>定金
              <input id="depositAmount" type="number" step="any" min="0" inputMode="decimal" placeholder="请输入定金" />
            </label>
          </div>
          <div className="shared-remark-panel" id="sharedRemarkPanel" hidden></div>
          <div className="row-list" id="rowList">
            <div className="empty-state">页面正在加载，如果手机上长时间停留在这里，请换用系统浏览器并刷新。</div>
          </div>
          <div className="bottom-row-actions">
            <button className="add-row-wide" id="addRowBottomBtn" type="button">+ 添加行</button>
            <button className="add-row-wide add-custom-wide" id="addCustomRowBottomBtn" type="button">+ 添加自由行</button>
          </div>
        </section>
      </section>

      <div className="natural-fill-dock" id="naturalFillDock">
        <div className="natural-fill-panel" id="naturalFillPanel" role="dialog" aria-modal="false" aria-labelledby="naturalFillTitle" hidden>
          <p className="natural-fill-prompt" id="naturalFillTitle">林峰帮你写：你可以告诉我窗帘的型号、安装位置等，我会帮你处理好表格</p>
          <label>描述内容
            <textarea id="naturalFillText" rows="4" placeholder="例如：客厅型号 A，宽 3.2 高 2.6，轨道 2 条单价 50"></textarea>
          </label>
          <div className="natural-fill-actions">
            <button className="secondary-action" id="naturalFillCancelBtn" type="button">取消</button>
            <button className="primary-action" id="naturalFillSubmitBtn" type="button">填入表格</button>
          </div>
        </div>
        <button className="natural-fill-launcher" id="naturalFillBtn" type="button" aria-label="智能填写" title="智能填写" aria-expanded="false" aria-controls="naturalFillPanel">智能填写</button>
      </div>

      <div className="status" id="status" role="status" aria-live="polite"></div>
      <noscript>
        <div className="status status-static is-error">当前浏览器未启用 JavaScript，无法显示表格内容。</div>
      </noscript>
      <script type="module" src="/app.js?v=20260522-native"></script>
    </main>
    </>
  );
}
