---
---

# settings.json

```
{
    "workbench.editor.openPositioning": "last",

    // エクスプローラの階層インデント
    "workbench.tree.indent": 20,

    // アクティブなタブの背景色を設定
    "workbench.colorCustomizations": {
        "tab.activeBackground": "#000000",
    },

    // これがtrueだとファイルを開いた際に前のタブのファイルを消してしまう
    "workbench.editor.enablePreview": false,

    // テキストの折り返し設定
    // この設定を切り替えるショートカットはwindows:[Alt + Z]、mac:[Option + Z]
    "editor.wordWrap": "off",

    "editor.minimap.enabled": false,

    // エディタにgit logのコメントが表示されないようにする
    "editor.codeLens": false,

    "editor.tabSize": 4,

    // インデントを空白文字で表す
    "editor.insertSpaces": true,

    "editor.wordSeparators": "/\\()\"':,.;<>~!@#$%^&*|+=[]{}`?-",

    // txtファイルをマークダウン形式で開く
    "files.associations": {
        "*.txt": "markdown",
    },

    ////////////////////////////////////////
    // 拡張機能:pythonがインストール済であること
    // そうするとpython設定項目が使えるようになる
    // 但しflake8は、各python環境(pipenvなど)ごとにインストールされている必要がある
    // インストールされていなければダイアログが出て自動でインストールしてくれる
    // python実行環境の設定は各フォルダのsettings.jsonで設定する
    // 設定方法
    //  * 手動でsettings.jsonに書き込む
    //  * コマンドパレット -> pythonインタープリターを選択、で環境を選択することで自動で書き込まれる

    "python.linting.flake8Enabled": true,
    "python.linting.pylintEnabled": false,

    "[python]": {
        // 入力した行を自動でコード整形する
        "editor.formatOnType": true
    },

    "workbench.editorAssociations": {
        "*.ipynb": "jupyter-notebook"
    },

    "notebook.cellToolbarLocation": {
        "default": "right",
        "jupyter-notebook": "left"
    },

    ////////////////////////////////////////
    // 拡張機能:vimがインストール済であること
    // そうするとvim設定項目が使えるようになる

    "vim.useSystemClipboard": true,
    "vim.ignorecase": false,
    "vim.easymotion": true,

    // nnoremap設定
    "vim.normalModeKeyBindingsNonRecursive": [
        //{
        //    mac用設定
        //    <C-s>はなぜかうまくバインドできない為、vscodeのキーボードショートカットでworkbench.action.files.saveを<C-s>にして対応した
        //    "before": ["<C-s>"],
        //    "after": [],
        //    "commands": [
        //        {"command": ":w"}
        //    ]   
        //},

        // C-rでファイルを閉じる
        {
            "before": ["<C-r>"],
            "after": [],
            "commands": [
                {"command": ":bd"}
            ]   
        },

        // C-jで行上移動(5行単位)
        {
            "before": ["<C-j>"],
            "after": ["5", "j"]
        },

        // C-jで行下移動(5行単位)
        {
            "before": ["<C-k>"],
            "after": ["5", "k"]
        },

        // jで折り返し行を見た目通りに移動
        {
            "before": ["j"],
            "after": ["g", "j"]
        },

        // kで折り返し行を見た目通りに移動
        {
            "before": ["k"],
            "after": ["g", "k"]
        },

        // f押下でeasy-motion起動
        {
            "before": ["f"],
            "after": ["leader", "leader", "leader", "b", "d", "w"]   
        },
    ],

    // inoremap設定
    "vim.insertModeKeyBindings": [
        // ;;でノーマルモードに戻るようにする。
        // 使いやすくて、できるだけ入力する機会が少ないキーとして;を割り当てた。
        // 但し、;を入力して補完などの機能を使う場合、;を入力して若干反応を待たなければならない。
        // まあ;を補完したい場合はほぼないと思うので問題はないか。
        {
            "before": [";", ";"],
            "after": ["<Esc>"]
        },

        {
            "before": ["<C-h>"],
            "after": ["<Left>"]
        },
    ],
    "git.openRepositoryInParentFolders": "never",
    "security.workspace.trust.untrustedFiles": "open",
    "editor.accessibilitySupport": "off",
    "workbench.editor.empty.hint": "hidden",
    "claudeCode.preferredLocation": "panel"
}
```

# keybindings.json

```
// 既定値を上書きするには、このファイル内にキー バインドを挿入します
[
    ////
    // エクスプローラへフォーカス
    ////
    {
        "key": "ctrl+e",
        "command": "workbench.view.explorer"
    },
    {
        "key": "ctrl+shift+e",
        "command": "-workbench.view.explorer"
    },
    ////
    // QuickOpen.InViewPicker(Ctrl+Q)中の上下移動
    ////
    {
        "key": "ctrl+n",
        "command": "workbench.action.quickOpenNavigateNextInViewPicker",
        "when": "inQuickOpen && inViewsPicker"
    },
    {
        "key": "ctrl+q",
        "command": "-workbench.action.quickOpenNavigateNextInViewPicker",
        "when": "inQuickOpen && inViewsPicker"
    },
    {
        "key": "ctrl+p",
        "command": "workbench.action.quickOpenNavigatePreviousInViewPicker",
        "when": "inQuickOpen && inViewsPicker"
    },
    {
        "key": "ctrl+shift+q",
        "command": "-workbench.action.quickOpenNavigatePreviousInViewPicker",
        "when": "inQuickOpen && inViewsPicker"
    },
    //// 
    // QuickOpen.InFilePicker(Ctrl+p)を(Ctrl+l)へ
    ////
    {
        "key": "ctrl+l",
        "command": "workbench.action.quickOpen"
    },
    {
        "key": "ctrl+p",
        "command": "-workbench.action.quickOpen"
    },
    ////
    // QuickOpen.InFilePicker(Ctrl+l)中の上下移動
    ////
    {
        "key": "ctrl+n",
        "command": "workbench.action.quickOpenNavigateNextInFilePicker",
        "when": "inFilesPicker && inQuickOpen"
    },
    {
        "key": "ctrl+p",
        "command": "-workbench.action.quickOpenNavigateNextInFilePicker",
        "when": "inFilesPicker && inQuickOpen"
    },
    {
        "key": "ctrl+p",
        "command": "workbench.action.quickOpenNavigatePreviousInFilePicker",
        "when": "inFilesPicker && inQuickOpen"
    },
    {
        "key": "ctrl+shift+p",
        "command": "-workbench.action.quickOpenNavigatePreviousInFilePicker",
        "when": "inFilesPicker && inQuickOpen"
    },
    ////
    // タブ移動を(Ctrl+n, Ctrl+p)に指定
    ////
    {
        "key": "ctrl+p",
        "command": "workbench.action.previousEditor"
    },
    {
        "key": "ctrl+pageup",
        "command": "-workbench.action.previousEditor"
    },
    {
        "key": "ctrl+n",
        "command": "workbench.action.nextEditor"
    },
    {
        "key": "ctrl+pagedown",
        "command": "-workbench.action.nextEditor"
    },
    {
        "key": "ctrl+f",
        "command": "actions.find",
        "when": "editorFocus || editorIsOpen"
    },
    {
        "key": "ctrl+f",
        "command": "-actions.find",
        "when": "editorFocus || editorIsOpen"
    },
    {
        "key": "ctrl+f",
        "command": "extension.vim_ctrl+f",
        "when": "editorTextFocus && vim.active && vim.use<C-f> && !inDebugRepl && vim.mode != 'Insert'"
    },
    {
        "key": "ctrl+f",
        "command": "-extension.vim_ctrl+f",
        "when": "editorTextFocus && vim.active && vim.use<C-f> && !inDebugRepl && vim.mode != 'Insert'"
    }
]
```
