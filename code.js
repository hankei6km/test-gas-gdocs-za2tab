/**
 * @param {DocumentApp.Document} doc
 * @param {Date} now
 */
function refreshTabsByLinks(doc, now = new Date()) {
  const tabs = doc.getTabs().filter(t => (t.getType() === DocumentApp.TabType.DOCUMENT_TAB))
  for (const tab of tabs) {
    refreshTabByLinks_(doc, tab, "web", now)
  }
}

/**
 * @param {DocumentApp.Document} doc
 * @param {Date} now
 * @param {{firstTabTitle:string}} opts
 */
function refreshFirstTabByLinks(doc, now = new Date(), opts = { firstTabTitle: '最新' }) {
  const tabs = doc.getTabs().filter(t => (t.getType() === DocumentApp.TabType.DOCUMENT_TAB))
  const firstTab = tabs.findIndex(t => t.getTitle() === opts.firstTabTitle)
  if (firstTab >= 0) {
    refreshTabByLinks_(doc, tabs[firstTab], "web", now)
  }
}

/**
 * @param {DocumentApp.Document} doc
 */
function refreshTabsByDocuments(doc, now = new Date()) {
  const tabs = doc.getTabs().filter(t => (t.getType() === DocumentApp.TabType.DOCUMENT_TAB))
  for (const tab of tabs) {
    refreshTabByLinks_(doc, tab, "documents", now)
  }
}

/**
 * @param {DocumentApp.Document} doc
 * @param {Date} now
 * @param {{firstTabTitle:string}} opts
 */
function refreshFirstTabByDocuments(doc, now = new Date(), opts = { firstTabTitle: '最新' }) {
  const tabs = doc.getTabs().filter(t => (t.getType() === DocumentApp.TabType.DOCUMENT_TAB))
  const firstTab = tabs.findIndex(t => t.getTitle() === opts.firstTabTitle)
  if (firstTab >= 0) {
    refreshTabByLinks_(doc, tabs[firstTab], "documents", now)
  }
}


/**
 * @param {DocumentApp.Document} doc
 */
function refreshTabsBySpreadsheets(doc, now = new Date()) {
  const tabs = doc.getTabs().filter(t => (t.getType() === DocumentApp.TabType.DOCUMENT_TAB))
  for (const tab of tabs) {
    refreshTabByLinks_(doc, tab, 'spreadsheets', now)
  }
}

/**
 * @param {DocumentApp.Document} doc
 * @param {Date} now
 * @param {{firstTabTitle:string}} opts
 */
function refreshFirstTabBySpreadsheets(doc, now = new Date(), opts = { firstTabTitle: '最新' }) {
  const tabs = doc.getTabs().filter(t => (t.getType() === DocumentApp.TabType.DOCUMENT_TAB))
  const firstTab = tabs.findIndex(t => t.getTitle() === opts.firstTabTitle)
  if (firstTab >= 0) {
    refreshTabByLinks_(doc, tabs[firstTab], "spreadsheets", now)
  }
}


/**
 * @param {DocumentApp.Document} doc
 * @param {{targetTabName:string|RegExp}} opts
 */
function refreshTabsByFolderFiles(doc, now = new Date(), opts = { targetTabName: /^(最新|(\d+(つ前|個前)))$/ }) {
  const tabs = doc.getTabs().filter(t => {
    const title = t.getTitle()
    return t.getType() === DocumentApp.TabType.DOCUMENT_TAB
      && (typeof (opts.targetTabName) === 'string' ?
        title === opts.targetTabName :
        title.match(opts.targetTabName) !== null
      )
  })
  {
    const len = tabs.length
    let i = 0
    let sourceUrls = []
    let files = null
    for (; i < len; i++) {
      sourceUrls = getSourceUrls_(tabs[i], 'folder')
      if (sourceUrls.length > 0) {
        const s = sourceUrls[0].url.split('/')
        const q = sourceUrls[0].q
        const orderBy = sourceUrls[0].orderBy
        const [folderId] = s[s.length - 1].split('?')
        files = getImages_(folderId, q, orderBy)
        break
      }
    }
    if (files) {
      for (const file of files) {
        refreshTabByFolderFiles_(doc, tabs[i], i, file, now)
        i++
        if (i >= len) {
          files.return()
        }

      }
    }
  }
}

/**
 * @param {DocumentApp.Document} doc
 * @param {{targetTabName:string|RegExp}} opts
 */
function refreshTabsBySheetRows(doc, now = new Date(), opts = { targetTabName: /^(最新|(\d+(つ前|個前)))$/ }) {
  const tabs = doc.getTabs().filter(t => {
    const title = t.getTitle()
    return t.getType() === DocumentApp.TabType.DOCUMENT_TAB
      && (typeof (opts.targetTabName) === 'string' ?
        title === opts.targetTabName :
        title.match(opts.targetTabName) !== null
      )
  })
  {
    const len = tabs.length
    let i = 0
    let sourceUrls = []
    let rows = null
    for (; i < len; i++) {
      sourceUrls = getSourceUrls_(tabs[i], 'spreadsheets')
      if (sourceUrls.length > 0) {
        const ss = SpreadsheetApp.openByUrl(sourceUrls[0].url)
        const sheet = (() => {
          const gid = getIdFromUrl_(sourceUrls[0].url)
          if (gid === null) {
            return ss.getActiveSheet()
          }
          return ss.getSheetById(gid)
        })()
        rows = getRows_(ss, sheet)
        break
      }
    }
    if (rows) {
      for (const row of rows) {
        refreshTabByFolderFiles_(doc, tabs[i], i, row, now)
        i++
        if (i >= len) {
          rows.return()
        }

      }
    }
  }
}


/**
 * @param {DocumentApp.Document} doc
 * @param {{firstTabTitle:string}} opts
 */
function rotateTabs(doc, opts = { firstTabTitle: '最新' }) {

  const tabs = doc.getTabs().filter(t => (t.getType() === DocumentApp.TabType.DOCUMENT_TAB))
  const firstTab = tabs.findIndex(t => t.getTitle() === opts.firstTabTitle)
  if (firstTab >= 0) {
    const tabNum = tabs.length
    for (let i = tabNum - 1; i > firstTab; i--) {
      const fromBody = tabs[i - 1].asDocumentTab().getBody().copy()
      const toBody = tabs[i].asDocumentTab().getBody().clear()
      for (const p of fromBody.getParagraphs()) {
        const c = p.copy()
        if (c.getType() === DocumentApp.ElementType.LIST_ITEM) {
          toBody.appendListItem(c).setAttributes(c.getAttributes())
        } else {
          toBody.appendParagraph(c).setAttributes(c.getAttributes())
        }
      }
      if (toBody.getParagraphs().length > 0) {
        toBody.removeChild(toBody.getChild(0))
      }
    }
  }

}

/**
 * @param {DocumentApp.Document} fromBody
 * @param {DocumentApp.Document} toBody
 */
function rollbackBody_(fromBody, toBody) {
  const paragraphs = fromBody.getParagraphs()
  toBody.clear()
  for (const p of paragraphs) {
    if (p.getType() === DocumentApp.ElementType.LIST_ITEM) {
      toBody.appendListItem(p.copy())
    } else {
      toBody.appendParagraph(p.copy())
    }
  }
  if (toBody.getParagraphs().length > 0) {
    toBody.removeChild(toBody.getChild(0))
  }
}

/**
 * @param {DocumentApp.Body} toBody
 * @param {boolean} all
 */
function clearBody_(toBody, all = false) {
  const p = toBody.getParagraphs()
  toBody.appendParagraph('') // LIST_ITEM TABLE などが最後だと removeChild などがエラーになるので(理由は不明)

  const len = p.length
  let top = 0
  if (!all) { // フォルダーファイルの場合、最初のタブ以外は全削除するため、フラグを設けている。
    // 頭出しする(最初に出てくるリスト項目が指示なのでリスト項目を残すため)
    for (; top < len && p[top].getType() !== DocumentApp.ElementType.LIST_ITEM; top++) {
    }
    for (; top < len && p[top].getType() === DocumentApp.ElementType.LIST_ITEM; top++) {
    }
  }
  for (let i = len - 1; i >= top; i--) {
    let parent = p[i].getParent()
    if (parent !== null && parent.getType() !== DocumentApp.ElementType.BODY_SECTION) {
      // 階層になっている場合の対応(LIST_ITEMは階層でないらしい)
      while (parent !== null && parent.getType() !== DocumentApp.ElementType.BODY_SECTION) {
        if (parent.getType() === DocumentApp.ElementType.TABLE) { // とりあえず TABLE だけ
          if (parent.getParent() !== null) { // 階層を巡回して消してないので、何度も確認する。無駄。
            parent.removeFromParent()
          }
        }
        parent = parent.getParent()
      }
    } else {
      toBody.removeChild(p[i])
    }
  }
}

/**
 * @param {DocumentApp.Document} doc
 * @param {DocumentApp.Tab} tab
 * @param {'web'|'folder'|'documents'|'spreadsheets'} kind
 * @param {Date} now
 */
function refreshTabByLinks_(doc, tab, kind = "web", now = (new Date())) {
  const body = tab.asDocumentTab().getBody()

  const sourceUrls = getSourceUrls_(tab, kind)
  if (sourceUrls.length > 0) {

    const fromBody = body.copy()
    const toBody = body
    clearBody_(toBody)
    try {
      toBody.appendParagraph(`${now.toLocaleString()} に取得`).setHeading(DocumentApp.ParagraphHeading.HEADING1)

      const ite = ((kind) => {
        if (kind === 'documents') {
          return getDocuments_(sourceUrls)
        } else if (kind === 'spreadsheets') {
          return getSpreadsheets_(sourceUrls)
        }
        return getFeeds_(sourceUrls)
      })(kind)

      for (const i of ite) {
        i.append(toBody)
      }
    } catch (e) {
      console.error('rollback')
      rollbackBody_(fromBody, toBody)
      throw (e)
    }
  }

}


/**
 * @param {DocumentApp.Document} doc
 * @param {DocumentApp.Tab} tab
 * @param {number} docTabIndex
 * @param  file
 * @param {Date} now
 */
function refreshTabByFolderFiles_(doc, tab, docTabIndex, file, now = (new Date())) {
  const body = tab.asDocumentTab().getBody()

  const fromBody = body.copy()
  const toBody = body
  /*const p = toBody.getParagraphs()
  const len = p.length
  toBody.appendParagraph('')
  for (let i = len - 1; i >= 0 && p[i].getType() !== DocumentApp.ElementType.LIST_ITEM; i--) {
    toBody.removeChild(p[i])
  }*/
  clearBody_(toBody, docTabIndex !== 0) // 最初の他部位以外は全削除
  try {
    file.append(toBody)
  } catch (e) {
    console.error('rollback')
    rollbackBody_(fromBody, toBody)
    throw (e)
  }
}


/**
 * @param {string} url
 * @returns {boolean}
 */
function isWebUrl_(url) {
  return !isFolderUrl_(url) && !isDocumentsUrl_(url) && !isSpreadsheetsUrl_(url)
}


/**
 * @param {string} url
 * @returns {boolean}
 */
function isFolderUrl_(url) {
  return url.startsWith('https://drive.google.com/drive/')
}

/**
 * @param {string} url
 * @returns {boolean}
 */
function isDocumentsUrl_(url) {
  return url.startsWith('https://docs.google.com/document/') || url.startsWith('https://docs.google.com/feeds/download/documents/export/Export')
}

/**
 * @param {string} url
 * @returns {boolean}
 */
function isSpreadsheetsUrl_(url) {
  return url.startsWith('https://docs.google.com/spreadsheets/d/')
}

/**
 * @param {DocumentApp.Tab} tab
 * @param {'web'|'folder'|'documents'|'spreadsheets'} kind
 */
function getSourceUrls_(tab, kind = 'web') {
  const ret = []
  const p = tab.asDocumentTab().getBody().getParagraphs()
  const len = p.length
  const list = []
  {
    let i = 0
    for (; i < len && p[i].getType() !== DocumentApp.ElementType.LIST_ITEM; i++) {
    }
    for (; i < len && p[i].getType() === DocumentApp.ElementType.LIST_ITEM; i++) {
      list.push(p[i])
    }
  }
  for (const item of list) {
    // const name = tab.getTitle()
    const text = item.getText()
    const url = item.getLinkUrl()
    if (typeof (text) === 'string' && text != '' && typeof (url) === 'string' && url !== '') {
      if ((kind === 'web' && isWebUrl_(url))
        || (kind === 'folder' && isFolderUrl_(url))
        || (kind === 'documents' && isDocumentsUrl_(url))
        || (kind === 'spreadsheets' && isSpreadsheetsUrl_(url))) {
        let name = text
        let q = ''
        let orderBy = ''
        if (kind === 'folder') {
          const t = text.split(' : ')
          name = t[0]
          q = t[1] || ''
          orderBy = t[2] || ''
        }
        ret.push({
          name,
          q,
          orderBy,
          url
        })
      }
    }
  }
  return ret
}

function splitReqs_(reqs, n) {
  const ret = []
  const len = reqs.length
  for (let i = 0; i < len; i += n) {
    ret.push(reqs.slice(i, i + n))
  }
  return ret;
}

/**
 * 
 * @returns {Generator<{append:(body:DocumentApp.Body)=>void},void,void>}
 */
function* getFeeds_(reqs) {
  for (const c of splitReqs_(reqs, 4)) {
    const r = UrlFetchApp.fetchAll(c.map(i => ({ url: i.url })))
    const len = r.length
    for (let i = 0; i < len; i++) {
      yield {
        append: ((res) => (body) => {
          console.log(res.name, res.content.length)
          body.appendParagraph(res.name).setHeading(DocumentApp.ParagraphHeading.HEADING2)
          body.appendParagraph(res.content)
          /*for (const text of textChunks_(i.content)) {
            toBody.appendParagraph(text)
          }*/
        })({
          content: r[i].getContentText(),
          ...c[i]
        }),
      }
    }
  }
}


/**
 * @param {string} url
 */
function urlAsExportAsMarkdown_(url) {
  const [_u, q] = url.split('?')
  if (q.match('exportFormat=markdown')) {
    return true
  }
  return false
}

/**
 * @param {string} url
 */
function urlAsExportAsCsv_(url) {
  const [_u, q] = url.split('?')
  if (q.match('format=csv')) {
    return true
  }
  return false
}


function exportDocumentsAsMd_(url) {
  // これはダメだった。
  // blob での操作が失敗する。ドキュメント的には 'text/markdown' は有効なはずなのだが(getContentType は成功する)。
  // Exception: Unexpected error while getting the method or property getDataAsString on object Blob.
  //const doc = DriveApp.getFileById(documentId)
  //const b = doc.getAs('text/markdown')

  // https://gist.github.com/tanaikech/0deba74c2003d997f67fb2b04dedb1d0
  //const url = `https://docs.google.com/feeds/download/documents/export/Export?exportFormat=markdown&id=${documentId}`;
  const res = UrlFetchApp.fetch(url, {
    headers: { authorization: "Bearer " + ScriptApp.getOAuthToken() },
  });
  return res.getBlob().getDataAsString()
}


function exportDocumentsAsCsv_(url) {
  // 上の AsMd と同じ。
  const res = UrlFetchApp.fetch(url, {
    headers: { authorization: "Bearer " + ScriptApp.getOAuthToken() }
  });
  return res.getBlob().getDataAsString().replace(/\r\n/g, '\n')
}


/**
 * 
 * @returns {Generator<{append:(body:DocumentApp.Body)=>void},void,void>}
 */
function* getDocuments_(reqs) {
  for (const c of reqs) {
    //const s = c.url.split('/')
    //const documentId = s[s.length - 2]
    yield {
      append: ((res) => (body) => {
        console.log(res.name)
        if (urlAsExportAsMarkdown_(res.url)) {
          const md = exportDocumentsAsMd_(res.url)
          body.appendParagraph(res.name).setHeading(DocumentApp.ParagraphHeading.HEADING2)
          body.appendParagraph(md)
        } else {
          const doc = DocumentApp.openByUrl(res.url)
          const tab = (() => {
            const tabId = getIdFromUrl_(res.url, keyName = 'tab')
            if (tabId === null) {
              return doc.getActiveTab()
            }
            return doc.getTab(tabId)
          })()

          const fromBody = tab.asDocumentTab().getBody()
          body.appendParagraph(res.name).setHeading(DocumentApp.ParagraphHeading.HEADING2)
          let i = 0;
          for (p of fromBody.getParagraphs()) {
            const c = p.copy()
            const attrs = c.getAttributes()
            const t = c.getType()
            try {
              if (t === DocumentApp.ElementType.LIST_ITEM) {
                body.appendListItem(c).setAttributes(attrs)
              } else {
                const a = body.appendParagraph(c).setAttributes(attrs)
                const heading = c.getHeading()
                if (heading === DocumentApp.ParagraphHeading.HEADING1) {
                  a.setHeading(DocumentApp.ParagraphHeading.HEADING3)
                } else if (heading === DocumentApp.ParagraphHeading.HEADING2) {
                  a.setHeading(DocumentApp.ParagraphHeading.HEADING4)
                } else if (heading === DocumentApp.ParagraphHeading.HEADING3) {
                  a.setHeading(DocumentApp.ParagraphHeading.HEADING5)
                } else if (heading === DocumentApp.ParagraphHeading.HEADING4) {
                  a.setHeading(DocumentApp.ParagraphHeading.HEADING6)
                }
              }
            } catch (e) {
              console.log(t.toString())
              throw (e)
            }
          }
          doc.saveAndClose()
        }
      })(
        {
          ...c
        }),
    }
  }
}

/**
 * @param {SpreadsheetApp.Sheet} sheet
 */
function getAppSheetDefaultImageFolderName_(sheet) {
  return `${sheet.getName()}_Images` // シート名に改行文字とかある？
}

/**
 * @param {SpreadsheetApp.Spreadsheet} ss
 * @param {SpreadsheetApp.Sheet} sheet
 * @returns {DriveApp.Folder | null}
 */
function getAppSheetDefaultImageFolder_(ss, folderName) {
  const parents = DriveApp.getFileById(ss.getId()).getParents()
  if (parents.hasNext()) {
    const parent = parents.next()
    if (!parents.hasNext()) { // 親は1つのときだけ(なんとなく、間違ったフォルダーを指す原因になりような気がしたので)
      let ret = null
      let folderCnt = 0
      const folders = parent.getFolders()
      while (folders.hasNext()) {
        const folder = folders.next()
        if (folder.getName() === folderName) {
          ret = folder
          folderCnt++
        }
      }
      if (folderCnt === 1) {
        return ret
      }
    }
    return null
  }
  return null
}

/**
 * @param {DriveApp.Folder} appSheetImageFolder
 * @param {string} appSheetImageFolderName
 * @param {string} value
 */
function cellValueToImageBlob_(appSheetImageFolder, appSheetImageFolderName, value) {
  if (value.startsWith('https://drive.google.com/open')) {
    // とりあえず
    const [_t, id] = value.split('id=')
    const file = DriveApp.getFileById(id)
    if (file.getMimeType().startsWith("image/")) {
      return file.getBlob()
    }
  } else if (appSheetImageFolder !== null && value.startsWith(appSheetImageFolderName)) {
    const pathParts = value.split('/')
    if (pathParts.length === 2) {// 階層が複数あるのはおかしい(AppSheet の仕様はわからないが、正しくないファイルを参照するような気もしないでもない)
      const files = appSheetImageFolder.getFilesByName(pathParts[1])
      if (files.hasNext()) {
        const file = files.next()
        if (!files.hasNext()) { // 複数存在するのはおかしい(AppSheet の仕様はわからないが、正しくないファイルを参照するような気もしないでもない)
          if (file.getMimeType().startsWith("image/")) {
            return file.getBlob()
          }
        }
      }
    }
  }
  return null
}


/**
 * @param {string} url
 * @param {string} keyName
 * @returns {string|null}
 */
function getIdFromUrl_(url, keyName = 'gid') {
  const [_t, qh] = url.split('?')
  if (typeof (qh) === 'string') {
    const [q, _h] = qh.split('#')
    if (typeof (q) === 'string') {
      const items = q.split('&')
      for (const item of items) {
        const [k, ...v] = item.split('=')
        if (k === keyName) {
          return v.join('')
        }
      }
    }
  }
  return null
}


/**
 * 
 * @returns {Generator<{append:(body:DocumentApp.Body)=>void},void,void>}
 */
function* getSpreadsheets_(reqs) {
  for (const c of reqs) {
    //const s = c.url.split('/')
    //const documentId = s[s.length - 2]
    yield {
      append: ((res) => (body) => {
        console.log(res.name)
        /*if (urlAsExportAsCsv_(res.url)) {
          const csv = exportDocumentsAsCsv_(res.url)
          body.appendParagraph(res.name).setHeading(DocumentApp.ParagraphHeading.HEADING2)
          body.appendParagraph(csv)
        }*/
        const ss = SpreadsheetApp.openByUrl(res.url)
        const sheet = (() => {
          const gid = getIdFromUrl_(res.url)
          if (gid === null) {
            return ss.getActiveSheet()
          }
          return ss.getSheetById(gid)
        })()
        console.log(sheet.getName())
        const range = sheet.getDataRange()
        const values = range.getValues()
        body.appendParagraph(res.name).setHeading(DocumentApp.ParagraphHeading.HEADING2)
        //body.appendParagraph(csv)
        // AppSheet 用
        const appSheetDefaultImageFolderName = getAppSheetDefaultImageFolderName_(sheet)
        const appSheetDefaultImageFolder = getAppSheetDefaultImageFolder_(ss, appSheetDefaultImageFolderName)
        const table = body.appendTable()
        for (const cells of values) {
          const tableRow = table.appendTableRow()
          for (const cell of cells) {
            const c = tableRow.appendTableCell()
            const imageBlob = cellValueToImageBlob_(appSheetDefaultImageFolder, appSheetDefaultImageFolderName, cell.toString())
            if (imageBlob === null) {
              c.setText(cell)
            } else {
              c.appendImage(imageBlob)
            }
          }
        }

      })(
        {
          ...c
        }),
    }
  }
}

const escapeQueryStringRegExp_ = new RegExp("'", 'g')
function escapeQueryString_(str) {
  return str.replace(escapeQueryStringRegExp_, "\\'")
}


/**
 * @param {string} folder
 * @param {string} q
 * @param {string} orderBy
 * @returns {Generator<{append:(body:DocumentApp.Body)=>void},void,void>}
 */
function* getImages_(folderId, q = '', orderBy = '') {
  let pageToken = undefined

  do {
    const f = Drive.Files.list({
      q: q ? `'${folderId}' in parents and trashed=false and (${q})` : `'${folderId}' in parents and trashed=false`,
      orderBy: orderBy || 'modifiedTime desc',
      pageToken
    })
    for (const file of f.files) {
      if (file.mimeType.startsWith("image/")) {
        console.log(file.name)
        yield {
          append: ((file) => (body) => {
            body.appendParagraph(file.getName()).setHeading(DocumentApp.ParagraphHeading.HEADING2)
            body.appendParagraph(`作成: ${file.getDateCreated().toLocaleString()}`)
            body.appendParagraph(`更新: ${file.getLastUpdated().toLocaleString()}`)
            body.appendImage(file.getBlob())
          })(DriveApp.getFileById(file.id)),
        }
      } else if (file.mimeType === 'application/pdf') {
        console.log(file.name)
        yield {
          append: ((file) => (body) => {
            body.appendParagraph(file.getName()).setHeading(DocumentApp.ParagraphHeading.HEADING2)
            body.appendParagraph(`作成: ${file.getDateCreated().toLocaleString()}`)
            body.appendParagraph(`更新: ${file.getLastUpdated().toLocaleString()}`)
            body.appendImage(file.getThumbnail())
          })(DriveApp.getFileById(file.id)),
        }
      } else if (file.mimeType === 'application/vnd.google-apps.document') {
        console.log(file.name)
        yield {
          append: ((doc) => (body) => {
            //const doc = DocumentApp.openByUrl(res.url)
            const fromBody = doc.getBody()
            body.appendParagraph(doc.getName()).setHeading(DocumentApp.ParagraphHeading.HEADING2)
            let i = 0;
            for (p of fromBody.getParagraphs()) {
              const c = p.copy()
              const attrs = c.getAttributes()
              const t = c.getType()
              try {
                if (t === DocumentApp.ElementType.LIST_ITEM) {
                  body.appendListItem(c).setAttributes(attrs)
                } else {
                  const a = body.appendParagraph(c).setAttributes(attrs)
                  const heading = c.getHeading()
                  if (heading === DocumentApp.ParagraphHeading.HEADING1) {
                    a.setHeading(DocumentApp.ParagraphHeading.HEADING3)
                  } else if (heading === DocumentApp.ParagraphHeading.HEADING2) {
                    a.setHeading(DocumentApp.ParagraphHeading.HEADING4)
                  } else if (heading === DocumentApp.ParagraphHeading.HEADING3) {
                    a.setHeading(DocumentApp.ParagraphHeading.HEADING5)
                  } else if (heading === DocumentApp.ParagraphHeading.HEADING4) {
                    a.setHeading(DocumentApp.ParagraphHeading.HEADING6)
                  }
                }
              } catch (e) {
                console.log(t.toString())
                throw (e)
              }
            }
            doc.saveAndClose()
          })(
            DocumentApp.openById(file.id)
          ),
        }
      }
    }
    pageToken = f.nextPageToken
  } while (pageToken != undefined)
}

/**
 * @param {Object[]} row
 */
function isBlankRow_(row) {
  return row.every(cell => cell.toString() === '')
}

/**
 * @param {SpreadsheetApp.Spreadsheet} ss 
 * @param {SpreadsheetApp.Sheet} sheet 
 * @returns {Generator<{append:(body:DocumentApp.Body)=>void},void,void>}
 */
function* getRows_(ss, sheet) {
  const range = sheet.getDataRange()
  const values = range.getValues()
  // AppSheet 用
  const appSheetDefaultImageFolderName = getAppSheetDefaultImageFolderName_(sheet)
  const appSheetDefaultImageFolder = getAppSheetDefaultImageFolder_(ss, appSheetDefaultImageFolderName)
  if (values.length > 0) {
    const titleRow = values[0]
    for (const row of values.slice(1).reverse()) {
      if (!isBlankRow_(row)) {
        yield {
          append: ((row) => (body) => {
            body.appendParagraph(sheet.getName()).setHeading(DocumentApp.ParagraphHeading.HEADING2)
            let len = row.length
            for (let i = 0; i < len; i++) {

              body.appendParagraph(titleRow[i].toString()).setHeading(DocumentApp.ParagraphHeading.HEADING3)
              const cell = row[i]
              const imageBlob = cellValueToImageBlob_(appSheetDefaultImageFolder, appSheetDefaultImageFolderName, cell.toString())
              if (imageBlob === null) {
                body.appendParagraph(cell.toString())
              } else {
                body.appendImage(imageBlob)
              }
            }
          })(row),
        }
      }
    }
  }

}



/**
 * 大きい文字列だと append できない可能性を考慮(いまは使ってない、サイズ以外にエラーになる要因がありそう)
 */
function* textChunks_(text, chunkSize = 200000) {
  // Intl.Segmenter などがないので妥協版。
  const chars = [...text];
  for (let i = 0; i < chars.length; i += chunkSize) {
    yield chars.slice(i, i + chunkSize).join('');
  }
}


