"use strict";

function _createForOfIteratorHelper(o, allowArrayLike) { var it = typeof Symbol !== "undefined" && o[Symbol.iterator] || o["@@iterator"]; if (!it) { if (Array.isArray(o) || (it = _unsupportedIterableToArray(o)) || allowArrayLike && o && typeof o.length === "number") { if (it) o = it; var i = 0; var F = function F() {}; return { s: F, n: function n() { if (i >= o.length) return { done: true }; return { done: false, value: o[i++] }; }, e: function e(_e2) { throw _e2; }, f: F }; } throw new TypeError("Invalid attempt to iterate non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); } var normalCompletion = true, didErr = false, err; return { s: function s() { it = it.call(o); }, n: function n() { var step = it.next(); normalCompletion = step.done; return step; }, e: function e(_e3) { didErr = true; err = _e3; }, f: function f() { try { if (!normalCompletion && it.return != null) it.return(); } finally { if (didErr) throw err; } } }; }

function _slicedToArray(arr, i) { return _arrayWithHoles(arr) || _iterableToArrayLimit(arr, i) || _unsupportedIterableToArray(arr, i) || _nonIterableRest(); }

function _nonIterableRest() { throw new TypeError("Invalid attempt to destructure non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); }

function _unsupportedIterableToArray(o, minLen) { if (!o) return; if (typeof o === "string") return _arrayLikeToArray(o, minLen); var n = Object.prototype.toString.call(o).slice(8, -1); if (n === "Object" && o.constructor) n = o.constructor.name; if (n === "Map" || n === "Set") return Array.from(o); if (n === "Arguments" || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(n)) return _arrayLikeToArray(o, minLen); }

function _arrayLikeToArray(arr, len) { if (len == null || len > arr.length) len = arr.length; for (var i = 0, arr2 = new Array(len); i < len; i++) { arr2[i] = arr[i]; } return arr2; }

function _iterableToArrayLimit(arr, i) { var _i = arr == null ? null : typeof Symbol !== "undefined" && arr[Symbol.iterator] || arr["@@iterator"]; if (_i == null) return; var _arr = []; var _n = true; var _d = false; var _s, _e; try { for (_i = _i.call(arr); !(_n = (_s = _i.next()).done); _n = true) { _arr.push(_s.value); if (i && _arr.length === i) break; } } catch (err) { _d = true; _e = err; } finally { try { if (!_n && _i["return"] != null) _i["return"](); } finally { if (_d) throw _e; } } return _arr; }

function _arrayWithHoles(arr) { if (Array.isArray(arr)) return arr; }

function asyncGeneratorStep(gen, resolve, reject, _next, _throw, key, arg) { try { var info = gen[key](arg); var value = info.value; } catch (error) { reject(error); return; } if (info.done) { resolve(value); } else { Promise.resolve(value).then(_next, _throw); } }

function _asyncToGenerator(fn) { return function () { var self = this, args = arguments; return new Promise(function (resolve, reject) { var gen = fn.apply(self, args); function _next(value) { asyncGeneratorStep(gen, resolve, reject, _next, _throw, "next", value); } function _throw(err) { asyncGeneratorStep(gen, resolve, reject, _next, _throw, "throw", err); } _next(undefined); }); }; }

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

function _defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } }

function _createClass(Constructor, protoProps, staticProps) { if (protoProps) _defineProperties(Constructor.prototype, protoProps); if (staticProps) _defineProperties(Constructor, staticProps); Object.defineProperty(Constructor, "prototype", { writable: false }); return Constructor; }

function _defineProperty(obj, key, value) { if (key in obj) { Object.defineProperty(obj, key, { value: value, enumerable: true, configurable: true, writable: true }); } else { obj[key] = value; } return obj; }

var Jade = /*#__PURE__*/function () {
  function Jade() {
    _classCallCheck(this, Jade);
  }

  _createClass(Jade, null, [{
    key: "load_gist",
    value: // Class Properties
    // Class Methods exposed to be called by Jade Users
    function () {
      var _load_gist = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee(gist_id) {
        var response, data, _i, _Object$values, file;

        return regeneratorRuntime.wrap(function _callee$(_context) {
          while (1) {
            switch (_context.prev = _context.next) {
              case 0:
                _context.prev = 0;
                _context.next = 3;
                return fetch("https://api.github.com/gists/".concat(gist_id, "?").concat(new Date()));

              case 3:
                response = _context.sent;
                _context.next = 6;
                return response.json();

              case 6:
                data = _context.sent;

                for (_i = 0, _Object$values = Object.values(data.files); _i < _Object$values.length; _i++) {
                  file = _Object$values[_i];
                  //console.log("===================================")
                  //console.log(file.content)
                  //console.log("===================================")
                  Jade.incorporate_code(file.content);
                }

                auto_exec();
                Jade.incorporate_code("auto_exec=null");
                _context.next = 14;
                break;

              case 12:
                _context.prev = 12;
                _context.t0 = _context["catch"](0);

              case 14:
              case "end":
                return _context.stop();
            }
          }
        }, _callee, null, [[0, 12]]);
      }));

      function load_gist(_x) {
        return _load_gist.apply(this, arguments);
      }

      return load_gist;
    }()
  }, {
    key: "import_code_module",
    value: function () {
      var _import_code_module = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee2(url_or_gist_id) {
        var url, url_data, response, data, files, gist_json, _i2, _Object$values2, file, name, _url_data, x, _i3, _files, _file, editor;

        return regeneratorRuntime.wrap(function _callee2$(_context2) {
          while (1) {
            switch (_context2.prev = _context2.next) {
              case 0:
                Jade.settings.workbook.module_to_import = url_or_gist_id; //console.log("at import code mod", Jade.settings.workbook)

                Jade.hide_element("import-module");
                Jade.save_settings();

                if (url_or_gist_id) {
                  _context2.next = 5;
                  break;
                }

                return _context2.abrupt("return");

              case 5:
                url = null;

                if (url_or_gist_id.substr(0, 4) === "http") {
                  // check to see if GIST url
                  if (url_or_gist_id.toLowerCase().includes("gist.github.com")) {
                    url_data = url_or_gist_id.split("/");
                    url = 'https://api.github.com/gists/' + data[data.length - 1] + "?" + new Date();
                  } else {
                    url = url_or_gist_id;
                  }
                } else {
                  // this looks like a gist id.  we should probably check it
                  // sometime.  for now, let's just assume it is
                  url = 'https://api.github.com/gists/' + url_or_gist_id + "?" + new Date();
                } // now we have the URL to process


                _context2.prev = 7;
                _context2.next = 10;
                return fetch(url);

              case 10:
                response = _context2.sent;
                _context2.next = 13;
                return response.text();

              case 13:
                data = _context2.sent;
                _context2.next = 21;
                break;

              case 16:
                _context2.prev = 16;
                _context2.t0 = _context2["catch"](7);
                alert(_context2.t0.message, "Error Loading Module");
                console.log("Error fetching gist", _context2.t0);
                return _context2.abrupt("return");

              case 21:
                //console.log("data",data)  
                files = []; // let's see if we got back json

                _context2.prev = 22;
                gist_json = JSON.parse(data); // there was no error parsing the json, so this must be gist manifest
                // load the individual files

                _context2.prev = 24;

                //console.log("parsing json from gist")
                for (_i2 = 0, _Object$values2 = Object.values(gist_json.files); _i2 < _Object$values2.length; _i2++) {
                  file = _Object$values2[_i2];
                  files.push({
                    name: file.filename.split(".js")[0],
                    code: file.content
                  });
                }

                _context2.next = 32;
                break;

              case 28:
                _context2.prev = 28;
                _context2.t1 = _context2["catch"](24);
                alert(_context2.t1.message, "Error Parsing Gist");
                return _context2.abrupt("return");

              case 32:
                _context2.next = 39;
                break;

              case 34:
                _context2.prev = 34;
                _context2.t2 = _context2["catch"](22);
                //console.log("module is not json")
                // json was not valid, assume we have js
                // check to see if there is a comment that specifies a module name
                //
                name = null;

                if (data.includes("ace.module:")) {
                  try {
                    //console.log("found ace lable")
                    name = JSON.parse(data.split("ace.module:")[1].split("*/")[0]).name;
                  } catch (e) {
                    //console.log("ace label invalid")
                    _url_data = url.split("/");
                    name = _url_data[_url_data.length - 1];
                  }
                }

                if (!name) {
                  //either there was no comment to specify a name, or there was an error in reading it
                  // we won't overwrite a module unless it is named and the name is something other than "Code"
                  // so here, we are going to give it a number to make it unique
                  //console.log("no name for you")
                  x = 1;

                  while (!!tag(Jade.panel_label_to_panel_name("Code " + x))) {
                    x++;
                  }

                  files.push({
                    name: "Code " + x,
                    code: data
                  });
                }

              case 39:
                // now we should have files looking like this
                // files:[{name:module1,content:"function zeta(){..."}, {name:module2,content:"function beta(){..."}]
                // we need to add or update based on the name.
                for (_i3 = 0, _files = files; _i3 < _files.length; _i3++) {
                  _file = _files[_i3];

                  //console.log(file.name,Jade.panel_label_to_panel_name(file.name))
                  if (!!tag(Jade.panel_label_to_panel_name(_file.name + " Module"))) {
                    // a module with this name already exists,  update
                    //console.log("========= ready to update ============", file.name)
                    editor = ace.edit(Jade.panel_label_to_panel_name(_file.name) + "_module-content");
                    editor.setValue(_file.code); //console.log(editor.getValue())
                  } else {
                    // no module with this name exists, append
                    Jade.add_code_module(_file.name, _file.code);
                  }
                }

              case 40:
              case "end":
                return _context2.stop();
            }
          }
        }, _callee2, null, [[7, 16], [22, 34], [24, 28]]);
      }));

      function import_code_module(_x2) {
        return _import_code_module.apply(this, arguments);
      }

      return import_code_module;
    }()
  }, {
    key: "set_css",
    value: function set_css(user_css) {
      Jade.css_suffix = user_css;
    }
  }, {
    key: "add_library",
    value: function add_library(url) {
      // adds a JS library to the head section of the HTML sheet
      var library = document.createElement('script');
      library.setAttribute('src', url); //console.log("library",library)

      document.head.appendChild(library);
    }
  }, {
    key: "close_canvas",
    value: function close_canvas() {
      Jade.panel_stack.pop();
      Jade.show_panel(Jade.panel_stack.pop());
    }
  }, {
    key: "open_editor",
    value: function open_editor() {
      Jade.show_panel(Jade.code_panels[0]);
    }
  }, {
    key: "open_output",
    value: function open_output() {
      Jade.show_panel("panel_output");
    }
  }, {
    key: "open_automations",
    value: function open_automations(show_close_button) {
      Jade.show_automations(show_close_button);
    }
  }, {
    key: "reset",
    value: function reset() {
      Jade.show_panel("panel_home");
    } // static show_html(html){
    //     //A simple function that is mapped differntly for examples than for modules
    //     //this is the module mapping
    //     Jade.open_canvas("html", html)
    // }

  }, {
    key: "open_canvas",
    value: function open_canvas(panel_name, html, show_panel_close_button, style_name) {
      if (style_name) {
        Jade.set_style(style_name);
      }

      if (!tag(panel_name)) {
        Jade.build_panel(panel_name);
      }

      if (!Jade.panels.includes(panel_name)) {
        Jade.panels.push(panel_name);
      }

      Jade.show_panel(panel_name);

      if (html) {
        if (show_panel_close_button || show_panel_close_button === undefined) {
          tag(panel_name).innerHTML = Jade.panel_close_button(panel_name) + html;
        } else {
          tag(panel_name).innerHTML = html;
        }
      }
    }
  }, {
    key: "print",
    value: function print(data, heading) {
      //if(!header && )
      if (!tag("panel_output").lastChild.lastChild.firstChild.tagName && !heading) {
        //no output here, need a headdng
        heading = "";
      }

      if (heading) {
        // there is a header, so make a new block
        //console.log("at data")
        var div = document.createElement("div");
        div.className = "jade-output";
        var header = document.createElement("div");
        header.className = "jade-output-header";
        var d = new Date();
        var ampm = " am";
        var hours = d.getHours();

        if (hours > 11) {
          ampm = "pm";

          if (hours > 12) {
            hours = hours - 12;
          }
        }

        header.innerHTML = '<span class="jade-output-time">' + hours + ":" + ("0" + d.getMinutes()).slice(-2) + ":" + ("0" + d.getSeconds()).slice(-2) + ampm + "</span> " + heading + '<div class="jade-output-close"><i class="fas fa-times" style="color:white;margin-right:.3rem;cursor:pointer" onclick="this.parentNode.parentNode.parentNode.remove()"></div>';
        var body = document.createElement("div");
        body.className = "jade-output-body";
        body.innerHTML = '<div style="margin:0;font-family: monospace;">' + data.replaceAll("\n", "<br />") + "<br />" + "</div>";
        div.appendChild(header);
        div.appendChild(body);
        tag("panel_output").appendChild(div);
      } else {
        // no header provided, append to most recently added
        tag("panel_output").lastChild.lastChild.firstChild.innerHTML += data.replaceAll("\n", "<br />") + "<br />";
      }
    }
  }, {
    key: "open_examples",
    value: function open_examples() {
      var panel_name = "panel_examples";
      Jade.set_style();
      Jade.show_panel(panel_name);
    } // Class methods that still need work before shown to the public

  }, {
    key: "set_theme",
    value: function set_theme(theme_name) {
      Jade.set_style(theme_name);
    }
  }, {
    key: "list_themes",
    value: function list_themes() {
      for (var _i4 = 0, _Object$entries = Object.entries(Jade.settings.workbook.styles); _i4 < _Object$entries.length; _i4++) {//console.log(theme, url)

        var _Object$entries$_i = _slicedToArray(_Object$entries[_i4], 2),
            theme = _Object$entries$_i[0],
            url = _Object$entries$_i[1];
      }
    } // Class Methods NOT meant to be called by Jade Users
    // It's not really a problem if they do, we just don't
    // think they are useful and we don't document them.

  }, {
    key: "officeReady",
    value: function officeReady(info) {
      // invoked when the office addin infrastructure has loaded
      if (info.host === Office.HostType.Excel) {
        Excel.run( /*#__PURE__*/function () {
          var _ref = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee3(excel) {
            var xl_settings, _ace_options;

            return regeneratorRuntime.wrap(function _callee3$(_context3) {
              while (1) {
                switch (_context3.prev = _context3.next) {
                  case 0:
                    xl_settings = excel.workbook.settings.getItemOrNullObject("jade").load("value"); //  const code_module_ids_from_settings = excel.workbook.settings.getItemOrNullObject("code_module_ids").load("value");

                    _context3.next = 3;
                    return excel.sync();

                  case 3:
                    if (xl_settings.isNullObject) {
                      // no Jade.settings so let's configuire some defaults
                      Jade.settings.user = {
                        ace_options: (_ace_options = {
                          selectionStyle: "line",
                          highlightActiveLine: true,
                          highlightSelectedWord: true,
                          readOnly: false,
                          copyWithEmptySelection: false,
                          cursorStyle: "ace",
                          mergeUndoDeltas: true,
                          behavioursEnabled: true,
                          wrapBehavioursEnabled: true,
                          enableAutoIndent: true,
                          showLineNumbers: true,
                          hScrollBarAlwaysVisible: false,
                          vScrollBarAlwaysVisible: false,
                          highlightGutterLine: true,
                          animatedScroll: false,
                          showInvisibles: false,
                          showPrintMargin: false,
                          printMarginColumn: 80,
                          printMargin: 80,
                          fadeFoldWidgets: false,
                          showFoldWidgets: true,
                          displayIndentGuides: true,
                          showGutter: true,
                          fontSize: "14pt",
                          scrollPastEnd: 0,
                          theme: "ace/theme/tomorrow",
                          maxPixelHeight: 0,
                          useTextareaForIME: true,
                          scrollSpeed: 2,
                          dragDelay: 0,
                          dragEnabled: true,
                          focusTimeout: 0,
                          tooltipFollowsMouse: true,
                          firstLineNumber: 1,
                          overwrite: false,
                          newLineMode: "auto",
                          useWorker: false,
                          useSoftTabs: true,
                          navigateWithinSoftTabs: false,
                          tabSize: 2,
                          wrap: true,
                          indentedSoftWrap: true,
                          foldStyle: "markbegin",
                          mode: "ace/mode/javascript",
                          enableMultiselect: true,
                          enableBlockSelect: true
                        }, _defineProperty(_ace_options, "tabSize", 2), _defineProperty(_ace_options, "useSoftTabs", true), _ace_options)
                      };
                      Jade.settings.workbook = {
                        code_module_ids: [],
                        examples_gist_id: "904983747625c3fdc8dfa69e0aaa0f08",
                        styles: {
                          system: null,
                          none: "/*No Theme CSS Used*/",
                          mvp: "https://cdnjs.cloudflare.com/ajax/libs/mvp.css/1.8.0/mvp.css",
                          marx: "https://cdnjs.cloudflare.com/ajax/libs/marx/4.0.0/marx.min.css",
                          water: "https://cdn.jsdelivr.net/npm/water.css@2/out/water.css",
                          "dark water": "https://cdn.jsdelivr.net/npm/water.css@2/out/dark.css",
                          sajura: "https://unpkg.com/sakura.css/css/sakura.css",
                          tacit: "https://cdn.jsdelivr.net/gh/yegor256/tacit@gh-pages/tacit-css-1.5.5.min.css",
                          pure: "https://unpkg.com/purecss@2.0.6/build/pure-min.css",
                          picnic: "https://cdn.jsdelivr.net/npm/picnic",
                          wing: "https://unpkg.com/wingcss",
                          chota: "https://unpkg.com/chota@latest",
                          bootstrap: "https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css"
                        }
                      };
                    } else {
                      // if xl_settings a null object
                      //console.log("xl_settings",xl_settings.value)
                      Jade.settings.workbook = xl_settings.value.workbook;
                      Jade.settings.user = xl_settings.value.user;
                    } // if Jade.settings null object
                    //console.log("before start_me_up, Jade.settings", Jade.settings)


                    Jade.configure_settings();
                    Jade.start_me_up();

                  case 6:
                  case "end":
                    return _context3.stop();
                }
              }
            }, _callee3);
          }));

          return function (_x3) {
            return _ref.apply(this, arguments);
          };
        }());
      } else {
        document.getElementById("sideload-msg").style.display = "flex";
        document.getElementById("menu").style.display = "none";
      }
    }
  }, {
    key: "start_me_up",
    value: function start_me_up() {
      //console.log("at start_me_up")
      Jade.settings.workbook.styles.system = tag("head_style").innerText;
      Jade.panels.push("panel_home"); //load code from one gist if specified.  
      //console.log("about ot load")

      if (Jade.settings.workbook.load_gist_id) {
        //console.log("in if")
        Jade.load_gist(Jade.settings.workbook.load_gist_id);
      } // add event listner to "add code module" input


      tag("new-module-name").addEventListener("keyup", function (event) {
        event.preventDefault();

        if (event.key === 'Enter') {
          tag("module-add-button").click();
        }
      }); // add event listner to "import module" input

      tag("gist-url").addEventListener("keyup", function (event) {
        event.preventDefault();

        if (event.key === 'Enter') {
          tag("module-import-button").click();
        }
      }); // fit the editor to the windows on resize

      window.addEventListener('resize', function (event) {
        //console.log("hi")
        var _iterator = _createForOfIteratorHelper(Jade.code_panels),
            _step;

        try {
          for (_iterator.s(); !(_step = _iterator.n()).done;) {
            var panel_name = _step.value;
            tag(panel_name + "_editor-page").style.height = Jade.editor_height();
          }
        } catch (err) {
          _iterator.e(err);
        } finally {
          _iterator.f();
        }
      }, true);
      Jade.init_examples();
      Jade.init_output(); // ---------------- Initializing Code Editors -----------------------------

      if (Jade.settings.workbook.code_module_ids.length > 0) {
        // show the button to view code modules
        Jade.show_element("open-editor");
      } //console.log("at init_code_editors       Jade.settings.workbook.code_module_ids", Jade.settings.workbook.code_module_ids)


      Excel.run( /*#__PURE__*/function () {
        var _ref2 = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee4(excel) {
          var parser, _iterator2, _step2, code_module_id, xmlpart, xmlBlob, doc, module_name, module_code, options;

          return regeneratorRuntime.wrap(function _callee4$(_context4) {
            while (1) {
              switch (_context4.prev = _context4.next) {
                case 0:
                  parser = new DOMParser();
                  _iterator2 = _createForOfIteratorHelper(Jade.settings.workbook.code_module_ids);
                  _context4.prev = 2;

                  _iterator2.s();

                case 4:
                  if ((_step2 = _iterator2.n()).done) {
                    _context4.next = 17;
                    break;
                  }

                  code_module_id = _step2.value;
                  xmlpart = excel.workbook.customXmlParts.getItem(code_module_id);
                  xmlBlob = xmlpart.getXml();
                  _context4.next = 10;
                  return excel.sync();

                case 10:
                  doc = parser.parseFromString(xmlBlob.value, "application/xml");
                  module_name = doc.getElementsByTagName("name")[0].textContent;
                  module_code = atob(doc.getElementsByTagName("code")[0].textContent); //const settings=atob(doc.getElementsByTagName("settings")[0].textContent)// might want to rename

                  options = atob(doc.getElementsByTagName("options")[0].textContent); //console.log("just loaded module", module_name)
                  //console.log("Jade.settings2", JSON.parse(Jade.settings))
                  //console.log("options", options)
                  //console.log("options-parsed", JSON.parse(options))

                  Jade.add_code_editor(module_name, module_code, code_module_id, null, JSON.parse(options));

                case 15:
                  _context4.next = 4;
                  break;

                case 17:
                  _context4.next = 22;
                  break;

                case 19:
                  _context4.prev = 19;
                  _context4.t0 = _context4["catch"](2);

                  _iterator2.e(_context4.t0);

                case 22:
                  _context4.prev = 22;

                  _iterator2.f();

                  return _context4.finish(22);

                case 25:
                case "end":
                  return _context4.stop();
              }
            }
          }, _callee4, null, [[2, 19, 22, 25]]);
        }));

        return function (_x4) {
          return _ref2.apply(this, arguments);
        };
      }()); //console.log("end of  start_me_up, Jade.settings", Jade.settings)
    }
  }, {
    key: "configure_settings",
    value: function configure_settings() {
      //console.log("Jade.settings", Jade.settings)
      //if(!tag('settings-page').className.includes("hidden")){
      //tag('jade-theme').focus();
      //console.log("fontSize", Jade.settings.user.ace_options.fontSize)  
      tag("jade-font-size").value = Jade.settings.user.ace_options.fontSize.replace("pt", "");

      if (Jade.settings.user.ace_options.wrap === false) {
        tag("jade-word-wrap").value = "no-wrap";
      } else if (Jade.settings.user.ace_options.indentedSoftWrap) {
        tag("jade-word-wrap").value = "wrap-indented";
      } else {
        tag("jade-word-wrap").value = "wrap";
      } //console.log("theme", Jade.settings.user.ace_options.theme)
      //console.log("theme", Jade.settings.user.ace_options.theme.split("/")[2])


      tag("examples-gist-id").value = Jade.settings.workbook.examples_gist_id;
      tag("jade-theme").value = Jade.settings.user.ace_options.theme.split("/")[2];
      tag("jade-line-numbers").checked = Jade.settings.user.ace_options.showGutter;

      if (Jade.settings.workbook.load_gist_id) {
        tag("load-gist-id").value = Jade.settings.workbook.load_gist_id;
      } //}

    }
  }, {
    key: "save_settings",
    value: function () {
      var _save_settings = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee5() {
        return regeneratorRuntime.wrap(function _callee5$(_context5) {
          while (1) {
            switch (_context5.prev = _context5.next) {
              case 0:
                if (tag("examples-gist-id").value && Jade.settings.workbook.examples_gist_id !== tag("examples-gist-id").value) {
                  // different examples specified.  Need to rebuild
                  Jade.rebuild_examples(tag("examples-gist-id").value);
                }

                Jade.settings.workbook.examples_gist_id = tag("examples-gist-id").value;
                Jade.settings.workbook.load_gist_id = tag("load-gist-id").value;
                Jade.settings.user.ace_options.theme = "ace/theme/" + tag("jade-theme").value;
                Jade.settings.user.ace_options.fontSize = tag("jade-font-size").value + "pt";
                Jade.settings.user.ace_options.showGutter = tag("jade-line-numbers").checked;
                _context5.t0 = tag("jade-word-wrap").value;
                _context5.next = _context5.t0 === "wrap" ? 9 : _context5.t0 === "wrap-indented" ? 12 : 15;
                break;

              case 9:
                Jade.settings.user.ace_options.wrap = true;
                Jade.settings.user.ace_options.indentedSoftWrap = false;
                return _context5.abrupt("break", 16);

              case 12:
                Jade.settings.user.ace_options.wrap = true;
                Jade.settings.user.ace_options.indentedSoftWrap = true;
                return _context5.abrupt("break", 16);

              case 15:
                Jade.settings.user.ace_options.wrap = "off";

              case 16:
                //console.log(Jade.settings.user.ace_options)
                Jade.apply_editor_options(Jade.settings.user.ace_options);
                _context5.next = 19;
                return Jade.write_settings_to_workbook();

              case 19:
                Jade.hide_element('settings-page');

              case 20:
              case "end":
                return _context5.stop();
            }
          }
        }, _callee5);
      }));

      function save_settings() {
        return _save_settings.apply(this, arguments);
      }

      return save_settings;
    }()
  }, {
    key: "write_settings_to_workbook",
    value: function () {
      var _write_settings_to_workbook = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee7() {
        return regeneratorRuntime.wrap(function _callee7$(_context7) {
          while (1) {
            switch (_context7.prev = _context7.next) {
              case 0:
                _context7.next = 2;
                return Excel.run( /*#__PURE__*/function () {
                  var _ref3 = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee6(excel) {
                    var xl_settings;
                    return regeneratorRuntime.wrap(function _callee6$(_context6) {
                      while (1) {
                        switch (_context6.prev = _context6.next) {
                          case 0:
                            xl_settings = excel.workbook.settings;
                            xl_settings.add("jade", Jade.settings); // adds or sets the value

                            _context6.next = 4;
                            return excel.sync();

                          case 4:
                          case "end":
                            return _context6.stop();
                        }
                      }
                    }, _callee6);
                  }));

                  return function (_x5) {
                    return _ref3.apply(this, arguments);
                  };
                }());

              case 2:
              case "end":
                return _context7.stop();
            }
          }
        }, _callee7);
      }));

      function write_settings_to_workbook() {
        return _write_settings_to_workbook.apply(this, arguments);
      }

      return write_settings_to_workbook;
    }()
  }, {
    key: "apply_editor_options",
    value: function apply_editor_options(options) {
      var _iterator3 = _createForOfIteratorHelper(Jade.code_panels),
          _step3;

      try {
        for (_iterator3.s(); !(_step3 = _iterator3.n()).done;) {
          var panel_name = _step3.value;
          //console.log("updating options on ", panel_name)
          var editor = ace.edit(panel_name + "-content");
          editor.setOptions(options);
        }
      } catch (err) {
        _iterator3.e(err);
      } finally {
        _iterator3.f();
      }
    }
  }, {
    key: "submit_feedback",
    value: function () {
      var _submit_feedback = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee8() {
        var message, form_data, options, response, data;
        return regeneratorRuntime.wrap(function _callee8$(_context8) {
          while (1) {
            switch (_context8.prev = _context8.next) {
              case 0:
                // send feedback to the google form
                message = [];

                if (tag("fb-type").value === "") {
                  message.push("<li>You must indicate the <b>type</b> of your feedback.</li>");
                }

                if (tag("fb-text").value === "") {
                  message.push("<li>You must provide the <b>text</b> of your feedback.</li>");
                }

                if (tag("fb-platform").value === "") {
                  message.push("<li>You must provide the <b>platform</b> you are using.</li>");
                }

                if (!(message.length === 0)) {
                  _context8.next = 51;
                  break;
                }

                // ready to submit
                form_data = [];
                form_data.push("entry.1033992853=");
                form_data.push(encodeURIComponent(tag("fb-type").value));
                form_data.push("&entry.482647522=");
                form_data.push(encodeURIComponent(tag("fb-platform").value));
                form_data.push("&entry.1230850697=");
                form_data.push(encodeURIComponent(tag("fb-email").value));
                form_data.push("&entry.1009690762=");
                form_data.push(encodeURIComponent(tag("fb-text").value));
                form_data.push("&entry.64124153=");
                form_data.push(encodeURIComponent("Addin"));
                tag("fb-message").innerHTML = "";
                options = {
                  method: 'POST',
                  mode: 'no-cors',
                  headers: {
                    'Content-Type': 'application/x-www-form-urlencoded'
                  },
                  body: form_data.join("")
                }; //console.log("form data",form_data.join(""))

                _context8.prev = 18;
                _context8.next = 21;
                return fetch("https://docs.google.com/forms/u/0/d/e/1FAIpQLSeRCZKHMU0HvXEiia8dn6hL0DfGoAH-PCJZolfiN1S_O90eSQ/formResponse", options);

              case 21:
                response = _context8.sent;
                _context8.next = 24;
                return response.text();

              case 24:
                data = _context8.sent;
                //console.log("data",data)
                tag("fb-message").style.color = "green"; //console.log("=====================",tag("fb-type").value)

                _context8.t0 = tag("fb-type").value;
                _context8.next = _context8.t0 === "Feature Resuest" ? 29 : _context8.t0 === "Report Issue" ? 31 : _context8.t0 === "Offer to help with development" ? 33 : _context8.t0 === "Praise for Addin" ? 35 : 37;
                break;

              case 29:
                message.push("Thanks for your feedback. Were not sure when we'll be updaing next but thanks for helping us understand your needs.");
                return _context8.abrupt("break", 38);

              case 31:
                message.push("Thanks for reporing this issue. While we cannot respond to every problem report, we'll do what we can.");
                return _context8.abrupt("break", 38);

              case 33:
                message.push("Thanks for offering to help on this project.  We'll take a look at your comment and get back to you--if you provided a valid email address.");
                return _context8.abrupt("break", 38);

              case 35:
                message.push("Thanks a million. We thrive on positive feedback!");
                return _context8.abrupt("break", 38);

              case 37:
                message.push("Thanks for asking this question. While we cannot respond to every question, we'll do what we can.");

              case 38:
                tag("fb-message").innerHTML = message.join("");
                tag("fb-message").scrollIntoView(true);
                setTimeout(function () {
                  Jade.hide_element('survey');
                  tag("fb-message").innerHTML = "";
                }, 10000);
                _context8.next = 49;
                break;

              case 43:
                _context8.prev = 43;
                _context8.t1 = _context8["catch"](18);
                //console.log("form error: ",e)
                tag("fb-message").style.color = "red";
                tag("fb-message").innerHTML = 'Oops.  It looks as though there was a network error.  Your can tray again later or submit at our <a href="https://docs.google.com/forms/d/e/1FAIpQLSeRCZKHMU0HvXEiia8dn6hL0DfGoAH-PCJZolfiN1S_O90eSQ/viewform?usp=pp_url&entry.64124153=web" target="_blank">Google Form<a>.';
                window.scrollTo(0, document.body.scrollHeight);
                tag("fb-message").scrollIntoView(true);

              case 49:
                _context8.next = 54;
                break;

              case 51:
                tag("fb-message").style.color = "red";
                tag("fb-message").innerHTML = "<ul>" + message.join("") + "</ul>";
                tag("fb-message").scrollIntoView(true);

              case 54:
              case "end":
                return _context8.stop();
            }
          }
        }, _callee8, null, [[18, 43]]);
      }));

      function submit_feedback() {
        return _submit_feedback.apply(this, arguments);
      }

      return submit_feedback;
    }()
  }, {
    key: "add_code_module",
    value: function add_code_module(name, code) {
      // a module built with whatever code is in Jade.default_code
      if (!name) {
        name = "code";
      } // check for duplicate name--that wreaks havoc


      var found_panel = false;

      var _iterator4 = _createForOfIteratorHelper(Jade.code_panels),
          _step4;

      try {
        for (_iterator4.s(); !(_step4 = _iterator4.n()).done;) {
          var panel_name = _step4.value;

          //console.log("panel_name",panel_name,Jade.panel_label_to_panel_name(name))
          if (panel_name === Jade.panel_label_to_panel_name(name) + "_module") {
            //we have a match, and that's a no-no
            alert('A module named "' + name + '" already exists in this workbook.  <br><br>Choose a differnt name.', "Invalid Module Name");
            return;
          }
        }
      } catch (err) {
        _iterator4.e(err);
      } finally {
        _iterator4.f();
      }

      if (!code) {
        // no code is pased in, determine which default code to import
        if (Jade.code_panels.length === 0) {
          code = Jade.default_code();
        } else {
          code = Jade.default_code("panel_" + name.toLowerCase().split(" ").join("_") + "_module");
        }
      }

      Jade.add_code_editor(name, code, "");
      Jade.hide_element("add-module");
      Jade.show_element("open-editor");
      Jade.show_panel(Jade.code_panels[Jade.code_panels.length - 1]);
      Jade.write_module_to_workbook(code, Jade.code_panels[Jade.code_panels.length - 1]);
    }
  }, {
    key: "get_style",
    value: function () {
      var _get_style = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee9(style_name, url, integrate_now) {
        var response, data;
        return regeneratorRuntime.wrap(function _callee9$(_context9) {
          while (1) {
            switch (_context9.prev = _context9.next) {
              case 0:
                if (integrate_now === undefined) {
                  integrate_now = true;
                }

                if (!Jade.settings.workbook.styles[style_name]) {
                  Jade.settings.workbook.styles[style_name] = url;
                }

                if (!(Jade.settings.workbook.styles[style_name].substr(0, 8) === "https://")) {
                  _context9.next = 11;
                  break;
                }

                _context9.next = 5;
                return fetch(Jade.settings.workbook.styles[style_name]);

              case 5:
                response = _context9.sent;
                _context9.next = 8;
                return response.text();

              case 8:
                data = _context9.sent;
                //console.log("data",data)
                Jade.settings.workbook.styles[style_name] = data;

                if (integrate_now) {
                  document.getElementById("head_style").remove();
                  document.head.insertAdjacentHTML("beforeend", '<style id="head_style" data-name="' + style_name + '">' + Jade.settings.workbook.styles[style_name] + Jade.css_suffix + "</style>");
                }

              case 11:
              case "end":
                return _context9.stop();
            }
          }
        }, _callee9);
      }));

      function get_style(_x6, _x7, _x8) {
        return _get_style.apply(this, arguments);
      }

      return get_style;
    }()
  }, {
    key: "set_style",
    value: function set_style(style_name) {
      //console.log("at set style", style_name)
      var css_sfx = Jade.css_suffix;

      if (!style_name) {
        style_name = "system";
        css_sfx = "";
      }

      if (Jade.settings.workbook.styles[style_name].substr(0, 8) === "https://") {
        // this style has not been fetched.  Get it now
        //console.log("in iff")
        Jade.get_style(style_name);
      }

      var style_tag = document.getElementById("head_style");

      if (style_tag.dataset.name !== style_name) {
        // only update the style tag if it is a differnt name
        style_tag.remove();
        document.head.insertAdjacentHTML("beforeend", '<style id="head_style" data-name="' + style_name + '">' + Jade.settings.workbook.styles[style_name] + css_sfx + "</style>");
      }
    }
  }, {
    key: "panel_close_button",
    value: function panel_close_button(panel_name) {
      return '<div id="close_' + panel_name + '" onclick="Jade.close_canvas(\'' + panel_name + '\')" class="top-corner" style="padding:5px 5px 0 5px;margin:5px 15px 0 0; cursor:pointer"><i class="far fa-window-close fa-2x"></i></i></div>';
    }
  }, {
    key: "build_panel",
    value: function build_panel(panel_name, show_close_button) {
      var div = document.createElement("div");
      div.id = panel_name;
      div.style.display = "none";
      div.innerHTML = Jade.panel_close_button(panel_name);

      if (!Jade.panels.includes(panel_name)) {
        Jade.panels.push(panel_name);
      }

      document.body.appendChild(div);

      if (show_close_button === undefined) {
        show_close_button = true;
      }

      if (!show_close_button) {
        Jade.hide_element("close_" + panel_name);
      }
    }
  }, {
    key: "show_automations",
    value: function show_automations(show_close_button) {
      var panel_name = "panel_listings"; // get the list of functions
      //###################################################### need to iterate over all modules

      var html = ['<h2 style="margin:0 0 0 1rem">Active Automations</h2><ol>']; //console.log("code panels", Jade.code_panels)

      var _iterator5 = _createForOfIteratorHelper(Jade.code_panels),
          _step5;

      try {
        for (_iterator5.s(); !(_step5 = _iterator5.n()).done;) {
          var code_panel = _step5.value;
          var editor = ace.edit(code_panel + "-content");
          var code = editor.getValue();
          var parsed_code = Jade.parse_code(code); //console.log(parsed_code)

          if (!parsed_code.error) {
            var _iterator6 = _createForOfIteratorHelper(parsed_code.body),
                _step6;

            try {
              for (_iterator6.s(); !(_step6 = _iterator6.n()).done;) {
                var element = _step6.value;
                var call_stmt = null;

                if (element.type === "FunctionDeclaration") {
                  if (element.id && element.id.name) {
                    // this is a named function
                    if (element.params.length === 0) {
                      //there are no params. it is callable
                      call_stmt = element.id.name + "()";
                    } else if (element.params.length === 1) {
                      // there is one param.  
                      if (element.async) {
                        if ("excel ctx context".includes(element.params[0].name)) {
                          // this is an async function with a single parameter named excel, ctx or context.  Run by passing context
                          call_stmt = "Excel.run(" + element.id.name + ")";
                        }
                      }
                    }

                    if (call_stmt) {
                      // this is a function we can run directly
                      // check for comment
                      var function_text = window[element.id.name] + ''; //console.log(function_text)

                      if (function_text.includes("Jade.listing:")) {
                        // this is a function we can run directly and it as the comment
                        //console.log("found a comment", func)
                        var comment = function_text.split("Jade.listing:")[1].split("*/")[0];

                        try {
                          var comment_json = JSON.parse(comment);
                          html.push('<li onclick="' + call_stmt + '" style="cursor:pointer"><b>' + comment_json.name + '</b>: ' + comment_json.description + '</li>');
                        } catch (e) {
                          ;
                          console.log("Jade.listing was not valid JSON", comment);
                        }
                      } //for function on code page

                    }
                  }
                }
              }
            } catch (err) {
              _iterator6.e(err);
            } finally {
              _iterator6.f();
            }
          }
        }
      } catch (err) {
        _iterator5.e(err);
      } finally {
        _iterator5.f();
      }

      if (html.length === 1) {
        //Did not find any properly configured functions
        html.push("There are currently no active automations in this workbook.");
      } else {
        html.push("</ul>");
      }

      Jade.open_canvas("panel_listings", html.join(""), show_close_button);
    }
  }, {
    key: "show_panel",
    value: function show_panel(panel_name) {
      if (Jade.code_panels.includes(panel_name)) {
        // set the size in case it is off
        if (tag(panel_name + "_function-names").length === 0) {
          // there are no function to run
          Jade.hide_element(panel_name + "_function-names");
        }

        tag(panel_name + "_editor-page").style.height = Jade.editor_height();

        try {
          ace.edit(panel_name + "-content").focus();
        } catch (e) {
          ;
          console.log("could not access ace.  This is expected", e);
        }
      } //################## 3 is  a problsm


      if (Jade.panels.slice(0, 3).includes(panel_name) || Jade.code_panels.includes(panel_name)) {
        Jade.set_style();
      } //console.log("trying",panel_name)


      var _iterator7 = _createForOfIteratorHelper(Jade.panels),
          _step7;

      try {
        for (_iterator7.s(); !(_step7 = _iterator7.n()).done;) {
          var panel = _step7.value;

          if (panel === panel_name) {
            //console.log("showing", panel)
            if (tag("selector_" + panel_name)) {
              tag("selector_" + panel_name).value = panel_name;
            }

            tag(panel).style.display = "block";
            Jade.panel_stack.push(panel);
          } else {
            //console.log(" hiding", panel)
            tag(panel).style.display = "none";
          }
        }
      } catch (err) {
        _iterator7.e(err);
      } finally {
        _iterator7.f();
      }

      if (Jade.code_panels.includes(panel_name)) {
        //focus the ace editor
        try {
          ace.edit(panel_name + "-content").focus();
        } catch (e) {
          ;
          console.log("could not access ace.  This is expected", e);
        }
      }
    }
  }, {
    key: "toggle_wrap",
    value: function toggle_wrap(panel_name) {
      var editor = ace.edit(panel_name + "-content");

      if (editor.getOptions().wrap === "off") {
        editor.setOptions({
          wrap: true
        });
        globalThis = true;
      } else {
        editor.setOptions({
          wrap: "off"
        });
      }
    }
  }, {
    key: "add_code_editor",
    value: function add_code_editor(module_name, code, module_xmlid, mod_settings, options_in) {
      // Jade.settings are things gove is storing with the module
      // options are the options from the ace editor
      //console.log("Jade.settings", Jade.settings)
      var options = Jade.settings.user.ace_options; //console.log(1)
      // not currently handling options at the editor level, so this block is diabled
      // if(options_in){// default options for the editor
      //   options=options_in
      // }
      //console.log(module_name, "options", options)

      if (!mod_settings) {
        mod_settings = {
          cursorPosition: {
            row: 0,
            column: 0
          }
        };
      } //console.log("adding ace editor", module_name, module_xmlid)


      var panel_name = "panel_" + module_name.toLowerCase().split(" ").join("_") + "_module";
      Jade.code_panels.push(panel_name);
      Jade.panel_labels.push(Jade.panel_name_to_panel_label(panel_name));
      Jade.panels.push(panel_name);
      Jade.build_panel(panel_name, false);
      tag(panel_name).dataset.module_name = module_name;
      tag(panel_name).dataset.module_xmlid = module_xmlid; //console.log(2)
      //console.log("initializing examples", tag(panel_name))

      tag(panel_name).appendChild(Jade.get_panel_selector(panel_name));
      var editor_container = document.createElement("div");
      editor_container.className = panel_name + "-content";
      var div = document.createElement("div");
      div.style.padding = ".2rem";
      div.style.verticalAlign = "middle";
      div.innerHTML = '<button title="Save code to workbook" onclick="Jade.update_editor_script(\'' + panel_name + '\')">Save</button> <button  title="Save code to workbook and execute" onclick="Jade.code_runner(tag(\'' + panel_name + '_function-names' + '\').value,\'' + panel_name + '\')">Run</button> <select id="' + panel_name + '_function-names"></select>';
      div.style.height = "22px";
      div.style.fontFamily = "auto";
      div.style.fontSize = "1rem";
      div.style.padding = ".2rem";
      div.style.backgroundColor = "#eee";
      div.id = panel_name + "_editor-bar";
      editor_container.appendChild(div); //console.log(3)
      //console.log("=======================================")
      //console.log(div);
      //console.log(div.clientHeight);

      var box = document.createElement("div");
      box.id = panel_name + "_editor-page";
      box.style.width = "100%";
      box.style.height = Jade.editor_height();
      box.style.display = "inline-block";
      box.style.position = "relative"; //console.log(4)
      //console.log("document",document.body.clientHeight);
      //console.log("scr",tag("panel_code_editor").Height);

      div = document.createElement("div");
      div.id = panel_name + "-content";
      div.dataset.edited = false;
      div.innerHTML = code.toHtmlEntities();
      box.appendChild(div); //console.log(5)

      editor_container.appendChild(box); //        elem.innerHTML = '<pre id="pre' + id + '">' + gist.script.content.split("<").join("&lt;") + '</pre>'

      var scriptPromise = new Promise(function (resolve, reject) {
        var script = document.createElement("script");
        document.body.appendChild(script);
        script.onload = resolve;
        script.onerror = reject;
        script.async = true;
        script.src = "https://cdnjs.cloudflare.com/ajax/libs/ace/1.4.13/ace.js";
      }); //console.log(6)

      scriptPromise.then(function () {
        var editor = ace.edit(panel_name + "-content");
        editor.on("blur", function () {
          window.event.target.parentElement.dataset.edited = true;
        }); //editor.setTheme("ace/theme/solarized_light");
        // if(!isNaN(options.fontSize)){
        //   options.fontSize += "pt"
        // }

        editor.setOptions(options);
        editor.session.setMode("ace/mode/javascript"); //console.log("Jade.settings", Jade.settings)

        editor.moveCursorTo(mod_settings.cursorPosition.row, mod_settings.cursorPosition.column);
        editor.commands.addCommand({
          // toggle word wrap
          name: "wrap",
          bindKey: {
            win: "Alt-z",
            mac: "Alt-z"
          },
          exec: function exec(editor) {
            var _iterator8 = _createForOfIteratorHelper(Jade.code_panels),
                _step8;

            try {
              for (_iterator8.s(); !(_step8 = _iterator8.n()).done;) {
                var panel = _step8.value;

                if (tag(panel).style.display === "block") {
                  // we found the one that is visible
                  Jade.toggle_wrap(panel);
                  break; //exit the loop
                }
              }
            } catch (err) {
              _iterator8.e(err);
            } finally {
              _iterator8.f();
            }
          }
        });
        editor.commands.addCommand({
          // could do ctrl+r but want to be parallel with save
          name: "run",
          bindKey: {
            win: "Ctrl-enter",
            mac: "Command-enter"
          },
          exec: function exec(editor) {
            var _iterator9 = _createForOfIteratorHelper(Jade.code_panels),
                _step9;

            try {
              for (_iterator9.s(); !(_step9 = _iterator9.n()).done;) {
                var panel = _step9.value;

                if (tag(panel).style.display === "block") {
                  // we found the one that is visible
                  Jade.code_runner(tag(panel + '_function-names').value, panel);
                  break; //exit the loop
                }
              }
            } catch (err) {
              _iterator9.e(err);
            } finally {
              _iterator9.f();
            }
          }
        });
        editor.commands.addCommand({
          // could do ctrl+r but want to be parallel with save
          name: "run_shift",
          bindKey: {
            win: "alt-enter",
            mac: "alt-enter"
          },
          exec: function exec(editor) {
            var _iterator10 = _createForOfIteratorHelper(Jade.code_panels),
                _step10;

            try {
              for (_iterator10.s(); !(_step10 = _iterator10.n()).done;) {
                var panel = _step10.value;

                if (tag(panel).style.display === "block") {
                  // we found the one that is visible
                  Jade.code_runner(tag(panel + '_function-names').value, panel);
                  break; //exit the loop
                }
              }
            } catch (err) {
              _iterator10.e(err);
            } finally {
              _iterator10.f();
            }
          }
        });
      });
      editor_container.style.display = "block";
      var parsed_code = Jade.parse_code(code);

      if (!parsed_code.error) {
        // only incorporate the code if it is free of syntax errors
        Jade.incorporate_code(code);
      }

      tag(panel_name).appendChild(editor_container);
      Jade.load_function_names_select(code, panel_name);

      if (mod_settings.func) {
        tag(panel_name + "_function-names").value = mod_settings.func;
      } // AutoExecutable function
      //console.log("about to autoexec")


      try {
        //console.log("in try")
        auto_exec();
      } catch (e) {
        ;
        console.log("catch", e);
      }
    }
  }, {
    key: "code_runner",
    value: function code_runner(script_name, panel_name) {
      //console.log(script_name,panel_name)
      if (tag(panel_name + "-content").dataset.edited === "true") {
        if (!Jade.update_editor_script(panel_name)) {
          // Jade.update_editor_script returns false if there is a 
          // syntax error.  Don't run the old code
          return;
        }
      }

      if (script_name.includes("(excel)")) {
        setTimeout("Excel.run(" + script_name.split("(")[0] + ")", 0); //run the function
      } else {
        setTimeout(script_name, 0); //run the function
      }
    }
  }, {
    key: "init_output",
    value: function init_output() {
      var panel_name = "panel_output";
      Jade.build_panel(panel_name, false);
      var panel = tag(panel_name); //console.log("initializing examples")

      panel.appendChild(Jade.get_panel_selector(panel_name));
      Jade.print('This panel shows the results of your calls to the "print" function.  Use "Jade.print(data)" to append text to the most recently printed block.', "About the Output Panel");
      Jade.print('\nUse "Jade.print(data, heading)" to start a new block.');
    }
  }, {
    key: "rebuild_examples",
    value: function rebuild_examples(gist_id) {
      tag("panel_examples").innerHTML = "";
      Jade.fill_examples(gist_id);
    }
  }, {
    key: "init_examples",
    value: function init_examples() {
      Jade.build_panel("panel_examples", false);
      Jade.fill_examples();
    }
  }, {
    key: "fill_examples",
    value: function fill_examples(gist_ids) {
      //gist_ids is a comma delimited list of gist ids that hold examples
      if (!gist_ids) {
        gist_ids = Jade.settings.workbook.examples_gist_id;
      }

      var panel_name = "panel_examples";
      var panel = tag(panel_name); //console.log("initializing examples")

      panel.appendChild(Jade.get_panel_selector(panel_name));
      var div = document.createElement("div");
      div.className = "content";
      div.id = "e_content";
      panel.appendChild(div); //console.log("gist_ids",gist_ids, Jade.settings)

      var gist_list = []; //console.log("gist_ids.split(",")",gist_ids.split(","))

      var _iterator11 = _createForOfIteratorHelper(gist_ids.split(",")),
          _step11;

      try {
        for (_iterator11.s(); !(_step11 = _iterator11.n()).done;) {
          var _gist = _step11.value;
          //console.log("gisting",gist)
          gist_list.push(_gist.trim());
        } //console.log("gist_list",gist_list)

      } catch (err) {
        _iterator11.e(err);
      } finally {
        _iterator11.f();
      }

      var html = [];

      for (var _i5 = 0, _gist_list = gist_list; _i5 < _gist_list.length; _i5++) {
        var gist = _gist_list[_i5];
        html.push("<div id=\"".concat(gist, "\"></div>"));
      }

      tag("e_content").innerHTML = html.join("");

      for (var i = 0; i < gist_list.length; i++) {
        //console.log("gist_list[i]",gist_list[i])
        Jade.get_example_html(gist_list[i], i + 1); // get the gist and integrate the examples
      }
    }
  }, {
    key: "get_example_html",
    value: function get_example_html(gist_id, sequence) {
      var url = "https://api.github.com/gists/".concat(gist_id, "?").concat(Date.now());
      var gist_url = "https://gist.github.com/".concat(gist_id); //console.log("building examples",url)
      //console.log("about to fetch",url)

      fetch(url).then(function (response) {
        return response.text();
      }).then(function (json_text) {
        // //console.log("json_text",json_text)
        var data = JSON.parse(json_text); // now we have the data from the api call.  need to organize it--especially for order
        // make a list of objects with integers as keys so the order will be numerical

        var filenames = {};

        for (var _i6 = 0, _Object$keys = Object.keys(data.files); _i6 < _Object$keys.length; _i6++) {
          var filename = _Object$keys[_i6];
          filenames[parseFloat(filename.split("_").shift())] = filename;
        } //console.log("filenames",filenames)


        var temp = data.description.split(":");
        var html = ["<h2 style=\"cursor:pointer\" title=\"Copy link to Example\" onclick=\"Jade.copy_to_clipboard('".concat(gist_url, "')\">").concat(temp.shift(), "</h2>")];
        html.push(temp.join(":").trim()); // iterate for each file in the gist

        var example_number = 0;

        for (var _i7 = 0, _Object$keys2 = Object.keys(filenames); _i7 < _Object$keys2.length; _i7++) {
          var key = _Object$keys2[_i7];
          example_number++; //console.log(key, filenames[key])

          temp = filenames[key].split(".");
          temp.pop(); //remove the extension

          temp = temp.join(".").split("_");
          temp.shift(); // remove the sort sequence number

          html.push("<p><a class=\"link\" title=\"Show Code\" onclick=\"Jade.show_example('".concat(sequence * 100 + example_number, "','").concat(data.files[filenames[key]].raw_url, "',").concat(data.files[filenames[key]].content.split(/\r\n|\r|\n/).length, ")\"><b>").concat(example_number, ". ").concat(temp.join(" "), ":</b> </a>"));
          temp = data.files[filenames[key]].content;

          if (temp.includes("*/")) {
            // there is a block comment.  assume it is a descriptions
            html.push(temp.split("*/")[0].split("/*")[1]);
          }

          html.push("</p><div id=\"page".concat(sequence * 100 + example_number, "\"></div>"));
        }

        tag(gist_id).innerHTML = html.join("");
        return;
      });
    }
  }, {
    key: "show_example",
    value: function show_example(id, url, lines) {
      // place the code in an editable box for user to see and play with
      // these examples should be made in script lab to have the right format
      var elem = tag("page" + id); //console.log(id + id);

      if (elem.innerHTML === "") {
        //console.log("lines", lines)
        elem.innerHTML = '<img id="loading-image" width="50" src="assets/loading.gif" />';
        fetch(url).then(function (response) {
          return response.text();
        }).then(function (data) {
          //console.log(data)
          //const gist = jsyaml.load(data);
          //console.log(gist)
          tag("loading-image").remove();
          var div = document.createElement("div");
          div.id = "example" + id + "_html"; //div.innerHTML = gist.template.content;

          div.style.marginBottom = "1rem";
          elem.appendChild(div);
          var box_height = lines * 22 + 17; // the size neede to show the whole example

          if (box_height > window.innerHeight) {
            box_height = window.innerHeight - 60;
          }

          var box = document.createElement("div");
          box.id = "page_" + id;
          box.style.width = "100%";
          box.style.height = box_height + "px";
          box.style.display = "inline-block";
          box.style.position = "relative";
          div = document.createElement("div");
          div.id = "editor" + id;
          div.innerHTML = data.toHtmlEntities();
          box.appendChild(div);
          elem.appendChild(box); // setting up the editor in the example space

          var scriptPromise = new Promise(function (resolve, reject) {
            var script = document.createElement("script");
            document.body.appendChild(script);
            script.onload = resolve;
            script.onerror = reject;
            script.async = true;
            script.src = "https://cdnjs.cloudflare.com/ajax/libs/ace/1.4.13/ace.js";
          });
          scriptPromise.then(function () {
            var editor = ace.edit("editor" + id);
            editor.on("blur", function () {
              Jade.update_script(id);
            });
            editor.setTheme("ace/theme/tomorrow"); //editor.session.$worker.send("changeOptions", [{ asi: true }]);

            editor.setOptions({
              fontSize: "14pt"
            });
            editor.getSession().setUseWorker(false);
            editor.session.setMode("ace/mode/javascript");
            editor.commands.addCommand({
              // toggle word wrap
              name: "wrap",
              bindKey: {
                win: "Alt-z",
                mac: "Alt-z"
              },
              exec: function exec(editor) {
                if (editor.getOptions().wrap === "off") {
                  editor.setOptions({
                    wrap: true
                  });
                  globalThis = true;
                } else {
                  editor.setOptions({
                    wrap: "off"
                  });
                }
              }
            });
          });
          elem.style.display = "block";
          Jade.incorporate_code(Jade.show_example_html_script(id));
          Jade.incorporate_code(data);
          setup(); // setup must be defined in the example

          if (!Jade.is_visible(tag("page_" + id))) {
            tag("page_" + id).scrollIntoView(false);
          }
        }).catch(function (error) {
          ;
          console.log(error);
        });
      } else {
        if (elem.style.display === "block") {
          elem.style.display = "none";
        } else {
          elem.style.display = "block";
          Jade.incorporate_code(Jade.show_example_html_script(id));
          Jade.incorporate_code(ace.edit("editor" + id).getValue());
          setup(); // setup must be defined in the example

          if (!Jade.is_visible(tag("page_" + id))) {
            tag("page_" + id).scrollIntoView(false);
          }
        }
      }
    }
  }, {
    key: "show_example_html_script",
    value: function show_example_html_script(id) {
      return 'function show_html(html){tag("example' + id + '_html").innerHTML=html}\n';
    }
  }, {
    key: "copy_to_clipboard",
    value: function copy_to_clipboard(text) {
      navigator.clipboard.writeText(text);
    }
  }, {
    key: "update_script",
    value: function update_script(id) {
      // read the script for an ace editor and write it to the DOM
      // this is the one used by the examples page
      //console.log("script" + id);
      var editor = ace.edit("editor" + id);
      var code = editor.getValue();
      var parsed_code = Jade.parse_code(code);

      if (parsed_code.error) {
        alert(parsed_code.error, "Syntax Error");
        editor.moveCursorTo(parsed_code.error.split("(")[1].split(":")[0] - 1, parsed_code.error.split("(")[1].split(":")[1].split(")")[0]);
        editor.focus();
        return false;
      }

      Jade.incorporate_code(Jade.show_example_html_script(id));
      Jade.incorporate_code(code);
      return true;
    }
  }, {
    key: "parse_code",
    value: function parse_code(code) {
      try {
        return acorn.parse(code, {
          "ecmaVersion": 8
        });
      } catch (e) {
        return {
          error: e.message
        };
      }
    }
  }, {
    key: "update_editor_script",
    value: function update_editor_script(panel_name) {
      // read the script for an ace editor and write it to the DOM
      // also saves the module to the custom properties
      //console.log("at Jade.update_editor_script", panel_name)
      // set the size of the editor in case there was a prior zoom
      // get the code
      var editor = ace.edit(panel_name + "-content");
      var code = editor.getValue(); // save the script to the workbook.  This is the most important thing
      // we are doing at the moment.  Do it first
      //console.log("about to write module")

      Jade.write_module_to_workbook(code, panel_name); // update the height of the editor in case it has gotten out of synch

      tag(panel_name + "_editor-page").style.height = Jade.editor_height(); // show_html is defined differntly for modules than for examples
      // we need to be sure it is defined correctly for modules, so we set it here

      Jade.incorporate_code('function show_html(html){open_canvas("html", html)}'); //Check for syntax errors

      var parsed_code = Jade.parse_code(code);

      if (parsed_code.error) {
        alert(parsed_code.error, "Syntax Error");
        editor.moveCursorTo(parsed_code.error.split("(")[1].split(":")[0] - 1, parsed_code.error.split("(")[1].split(":")[1].split(")")[0]);
        editor.focus();
        return false;
      } // load the user's code into the browser


      Jade.incorporate_code(code); // put the function names in the function dropdown

      Jade.load_function_names_select(parsed_code, panel_name);
      return true;
    }
  }, {
    key: "incorporate_code",
    value: function incorporate_code(code) {
      // It turns out that the script block does not need to persist in the HTML
      // once the script block is loaded, the JS is parsed and not again referenced.
      // So, we create a script block, append it to the document body, then remove.
      var script = document.createElement("script");
      script.innerHTML = code;
      document.body.appendChild(script);
      document.body.lastChild.remove();
    }
  }, {
    key: "write_module_to_workbook",
    value: function write_module_to_workbook(code, panel_name) {
      var options = {
        fontSize: "14pt"
      };
      var settings = {
        cursorPosition: {
          row: 0,
          column: 0
        },
        func: tag(panel_name + "_function-names").value
      };

      try {
        var editor = ace.edit(panel_name + "-content");
        Jade.settings.cursorPosition = editor.getCursorPosition();
        options = editor.getOptions();
      } catch (e) {//console.log("This is an expected error: ace editor not yet built",e)
      } //console.log("writing options", JSON.stringify(options))


      var module_name = tag(panel_name).dataset.module_name;
      var xmlid = tag(panel_name).dataset.module_xmlid;

      if (xmlid) {
        // workbook as already been saved and has and xmlid
        //console.log("saving an existing book")
        Jade.save_module_to_workbook(code, module_name, Jade.settings, xmlid, options);
      } else {
        // workbook not yet saved, function call will return an xmlid
        Jade.save_module_to_workbook(code, module_name, Jade.settings, xmlid, options, tag(panel_name));
      }
    }
  }, {
    key: "save_module_to_workbook",
    value: function save_module_to_workbook(code, module_name, mod_settings, xmlid, options, tag_to_hold_new_xml_id) {
      // options are currently being ignored, they are handled globally instead of at the module level
      // save to workbook
      Excel.run( /*#__PURE__*/function () {
        var _ref4 = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee10(excel) {
          var module_xml, customXmlPart, _customXmlPart;

          return regeneratorRuntime.wrap(function _callee10$(_context10) {
            while (1) {
              switch (_context10.prev = _context10.next) {
                case 0:
                  //console.log("saving code", panel_name) 
                  if (!mod_settings) {
                    mod_settings = {
                      cursorPosition: {
                        row: 0,
                        column: 0
                      }
                    };
                  } //The next line has been disabled because we are not currently maintaining options at the module level
                  //const module_xml = "<module xmlns='http://schemas.gove.net/code/1.0'><name>"+module_name+"</name><Jade.settings>"+btoa(JSON.stringify(Jade.settings))+"</Jade.settings><options>"+btoa(JSON.stringify(options))+"</options><code>"+btoa(code)+"</code></module>"
                  // save module without options


                  module_xml = "<module xmlns='http://schemas.gove.net/code/1.0'><name>" + module_name + "</name><Jade.settings>" + btoa(JSON.stringify(mod_settings)) + "</Jade.settings><options>" + btoa(null) + "</options><code>" + btoa(code) + "</code></module>";

                  if (!xmlid) {
                    _context10.next = 8;
                    break;
                  }

                  //console.log("updating xml", xmlid, typeof xmlid)
                  customXmlPart = excel.workbook.customXmlParts.getItem(xmlid);
                  customXmlPart.setXml(module_xml);
                  excel.sync(); //console.log("------- launched saving: existing -------")

                  _context10.next = 15;
                  break;

                case 8:
                  //console.log("creating xml")
                  _customXmlPart = excel.workbook.customXmlParts.add(module_xml);

                  _customXmlPart.load("id");

                  _context10.next = 12;
                  return excel.sync();

                case 12:
                  //console.log("customXmlPart",customXmlPart.getXml())
                  // this is a newly created module and needs to have a custom xmlid part made for it
                  //console.log("23443", Jade.settings, customXmlPart.id)
                  Jade.settings.workbook.code_module_ids.push(_customXmlPart.id); // add the id to the list of ids

                  Jade.write_settings_to_workbook(); //console.log("------- launched saving: newly created -------")
                  //console.log(typeof tag_to_hold_new_xml_id, tag_to_hold_new_xml_id)

                  if (tag_to_hold_new_xml_id) {
                    //console.log("writing.....xmlid", customXmlPart.id)
                    tag_to_hold_new_xml_id.dataset.module_xmlid = _customXmlPart.id; //console.log(tag_to_hold_new_xml_id)
                  }

                case 15:
                case "end":
                  return _context10.stop();
              }
            }
          }, _callee10);
        }));

        return function (_x9) {
          return _ref4.apply(this, arguments);
        };
      }());
    }
  }, {
    key: "load_function_names_select",
    value: function load_function_names_select(code, panel_name) {
      // reads the function names from the code and puts them in the function name select
      // if a string is passed in, parse it.  otherwise, assume it is already parsed
      if (typeof code === "string") {
        var parsed_code = Jade.parse_code(code);
      } else {
        var parsed_code = code;
      }

      if (parsed_code.error) {
        return;
      }

      var selectElement = tag(panel_name + "_function-names");
      var selected_script = selectElement.value;

      while (selectElement.options.length > 0) {
        selectElement.remove(0);
      }

      var _iterator12 = _createForOfIteratorHelper(parsed_code.body),
          _step12;

      try {
        for (_iterator12.s(); !(_step12 = _iterator12.n()).done;) {
          var element = _step12.value;
          var option_value = null;

          if (element.type === "FunctionDeclaration") {
            if (element.id && element.id.name) {
              // this is a named function
              if (element.params.length === 0) {
                //there are no params. it is callable
                option_value = element.id.name + "()";
              } else if (element.params.length === 1) {
                // there is one param.  
                if (element.async) {
                  if ("excel ctx context".includes(element.params[0].name)) {
                    // this is an async function with a single parameter named excel, ctx or context.  Run by passing context
                    option_value = element.id.name + "(excel)";
                  }
                }
              }

              if (option_value) {
                // this is a function we can run directly
                var option = document.createElement("option");

                if (option_value === selected_script) {
                  option.selected = 'selected';
                }

                option.text = element.id.name;
                option.value = option_value;
                selectElement.add(option);
              }
            }
          }
        }
      } catch (err) {
        _iterator12.e(err);
      } finally {
        _iterator12.f();
      }

      if (selectElement.length === 0) {
        Jade.hide_element(selectElement);
      } else {
        Jade.show_element(selectElement);
      }
    }
  }, {
    key: "get_panel_selector",
    value: function get_panel_selector(panel) {
      var panel_label = Jade.panel_name_to_panel_label(panel);
      var sel = document.createElement("select"); //console.log("appending panel=====", panel)

      if (!Jade.panel_labels.includes(panel_label)) {
        Jade.panel_labels.push(panel_label);
      }

      sel.className = "panel-selector"; // put the options in this panel selector

      Jade.update_panel_selector(sel); // update all others panel selectirs

      var _iterator13 = _createForOfIteratorHelper(document.getElementsByClassName("panel-selector")),
          _step13;

      try {
        for (_iterator13.s(); !(_step13 = _iterator13.n()).done;) {
          var selector = _step13.value;
          Jade.update_panel_selector(selector);
        }
      } catch (err) {
        _iterator13.e(err);
      } finally {
        _iterator13.f();
      }

      sel.value = panel;
      sel.style.height = "40px";
      sel.id = "selector_" + panel;
      sel.onchange = Jade.select_page;
      return sel;
    }
  }, {
    key: "update_panel_selector",
    value: function update_panel_selector(sel) {
      // put the proper choices in a panel selector
      while (sel.length > 0) {
        sel.remove(0);
      }

      for (var i = 0; i < Jade.panel_labels.length; i++) {
        var option = document.createElement("option");
        option.value = Jade.panel_label_to_panel_name(Jade.panel_labels[i]); //console.log("-->", option.value)

        option.text = Jade.panel_labels[i];
        option.className = "panel-selector-option";
        sel.appendChild(option);
      }
    }
  }, {
    key: "panel_label_to_panel_name",
    value: function panel_label_to_panel_name(panel_label) {
      return "panel_" + panel_label.toLocaleLowerCase().split(" ").join("_");
    }
  }, {
    key: "panel_name_to_panel_label",
    value: function panel_name_to_panel_label(panel_name) {
      var panel_label = panel_name.replace("panel_", "");
      panel_label = panel_label.split("_").join(" ");
      return panel_label.toTitleCase();
    }
  }, {
    key: "select_page",
    value: function select_page() {
      //console.log("panel name",window.event.target.value)
      Jade.show_panel(window.event.target.value);
    }
  }, {
    key: "show_element",
    value: function show_element(tag_id) {
      // removes the hidden class from a tag's css
      if (typeof tag_id === "string") {
        var the_tag = tag(tag_id);
      } else {
        the_tag = tag_id;
      }

      the_tag.className = the_tag.className.replaceAll("hidden", "");
    }
  }, {
    key: "hide_element",
    value: function hide_element(tag_id) {
      // adds the hidden class from a tag's css
      //console.log("tag_id",tag_id)
      //console.log("tag(tag_id)",tag(tag_id))
      //console.log("tag(tag_id).className",tag(tag_id).className)
      if (typeof tag_id === "string") {
        var the_tag = tag(tag_id);
      } else {
        the_tag = tag_id;
      }

      if (the_tag) {
        if (the_tag.className) {
          if (!the_tag.className.includes("hidden")) {
            the_tag.className = (the_tag.className + " hidden").trim();
          }
        } else {
          the_tag.className = "hidden";
        }
      }
    }
  }, {
    key: "toggle_element",
    value: function toggle_element(tag_id) {
      // adds the hidden class from a tag's css
      //console.log("tag_id", tag_id)
      if (typeof tag_id === "string") {
        var the_tag = tag(tag_id);
      } else {
        the_tag = tag_id;
      }

      if (the_tag.className.includes("hidden")) {
        Jade.show_element(the_tag);
      } else {
        Jade.hide_element(the_tag);
      }
    }
  }, {
    key: "editor_height",
    value: function editor_height() {
      return window.innerHeight - 73 + "px";
    }
  }, {
    key: "default_code",
    value: function default_code(panel_name) {
      var code = "async function write_timestamp(excel){\n      /*Jade.listing:{\"name\":\"Timestamp\",\"description\":\"This sample function records the current time in the selected cells\"}*/\n    excel.workbook.getSelectedRange().values = new Date();\n    await excel.sync();\n  }\n  \n  function auto_exec(){\n    // This function is called when the addin opens.\n    // un-comment a line below to take action on open.\n  \n    // Jade.open_automations() // displays a list of functions for a user\n  ";

      if (panel_name) {
        code += "  // Jade.show_panel('".concat(panel_name, "')      // shows this code editor\n  }");
      } else {
        code += "  // Jade.open_editor()      // shows the code editor\n  }";
      }

      return code;
    }
  }, {
    key: "show_import_module",
    value: function show_import_module() {
      //console.log("at show import", Jade.settings.workbook.module_to_import)
      if (Jade.settings.workbook.module_to_import) {
        //console.log("in if")
        tag("gist-url").value = Jade.settings.workbook.module_to_import;
      }

      Jade.toggle_element('import-module');
      tag('gist-url').focus();
    }
  }, {
    key: "is_visible",
    value: function is_visible(el) {
      var rect = el.getBoundingClientRect();
      return rect.top >= 0 && rect.left >= 0 && rect.bottom <= (window.innerHeight || document.documentElement.clientHeight) && rect.right <= (window.innerWidth || document.documentElement.clientWidth);
    }
  }]);

  return Jade;
}(); // end of Jade class

/*** Convert a string to HTML entities */


_defineProperty(Jade, "settings", {});

_defineProperty(Jade, "css_suffix", "");

_defineProperty(Jade, "panels", ['panel_home', 'panel_examples']);

_defineProperty(Jade, "panel_labels", ["Home", "Examples", "Output"]);

_defineProperty(Jade, "code_panels", []);

_defineProperty(Jade, "panel_stack", ['panel_home']);

String.prototype.toHtmlEntities = function () {
  return this.replace(/./gm, function (s) {
    // return "&#" + s.charCodeAt(0) + ";";
    return s.match(/[a-z0-9\s]+/i) ? s : "&#" + s.charCodeAt(0) + ";";
  });
};

String.prototype.toTitleCase = function () {
  str = this.toLowerCase().split(' ');

  for (var i = 0; i < str.length; i++) {
    str[i] = str[i].charAt(0).toUpperCase() + str[i].slice(1);
  }

  return str.join(' ');
};

function tag(id) {
  // a short way to get an element by ID
  return document.getElementById(id);
}

function alert(data, heading) {
  if (tag("jade-alert")) {
    tag("jade-alert").remove();
  }

  if (!heading) {
    heading = "System Message";
  }

  var div = document.createElement("div");
  div.className = "jade-alert";
  div.id = 'jade-alert';
  var header = document.createElement("div");
  header.className = "jade-alert-header";
  header.innerHTML = heading + '<div class="jade-alert-close"><i class="fas fa-times" style="color:white;margin-right:.3rem;cursor:pointer" onclick="this.parentNode.parentNode.parentNode.remove()"></div>';
  var body = document.createElement("div");
  body.className = "jade-alert-body";
  body.innerHTML = data;
  div.appendChild(header);
  div.appendChild(body);
  document.body.appendChild(div);
}