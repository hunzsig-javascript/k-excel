import React from 'react';
import {message, Modal, Alert, Progress, Button, Upload} from 'antd';
import {Api} from 'api';
import {I18n} from "basic";
import KExcel from './KExcel';

const styles = {
  excelBox: {
    backgroundColor: 'rgb(255, 255, 255)',
    borderRadius: '6px',
    padding: '20px',
    marginBottom: '20px',
  },
};
const Index = {
  sheetLength: [
    {value: 1, label: '1'},
    {value: 20, label: '20'},
    {value: 50, label: '50'},
    {value: 100, label: '100'},
    {value: 500, label: '500'},
    {value: 1000, label: '1000'},
  ],
  pullLoop: (scope, params, set, percent, result) => {
    result = result || [];
    params.page = 1;
    params.pagePer = 10;
    params.pageCurrent = params.pageCurrent || 0;
    Api.real(scope, params, (res) => {
      if (res.code === 200) {
        if (res.data.page.total > 0) {
          res.data.data.forEach((item) => {
            result.push(JSON.parse(JSON.stringify(item)));
          });
          if (res.data.page.current < res.data.page.end - 1) {
            params.pageCurrent += 1;
            Index.pullLoop(scope, params, set, percent, result);
            if (typeof percent === 'function') {
              percent(res.data.page.current / res.data.page.end);
            }
          } else {
            const toExcel = new KExcel();
            toExcel.excelZip(result, res.data.page, set);
            if (typeof percent === 'function') {
              percent(100);
            }
          }
        } else if (typeof percent === 'function') {
          percent(-1);
        }
      } else {
        Index.pullLoop(scope, params, set, percent, result);
      }
    });
  },
  pushLoop: (element, set, then) => {
    console.log(set);
    const pullExcel = new KExcel();
    pullExcel.excelPull(element, set, then);
  },
  pull: (opts) => {
    opts.tips = opts.tips || I18n('EXCEL_EXPORTING');
    opts.isZip = opts.isZip || 1;
    opts.filter = opts.filter || [];
    opts.sheet = opts.sheet || [];
    opts.scope = opts.scope || null;
    opts.form = null;
    if (!Array.isArray(opts.sheet) || opts.sheet.length <= 0) {
      message.error('not sheet');
      return;
    }
    if (!opts.scope) {
      message.error('not scope');
      return;
    }
    const id = (new Date()).getTime();
    const modal = Modal.warn({
      width: 700,
      title: I18n('EXCEL_EXPORT'),
      maskClosable: true,
      className: 'vertical-center-modal hideFooter',
      content: (
        <div>
          <Alert message={opts.tips} type="warning" showIcon={false} banner/>
          <ThisForm form={{
            refresh: true,
            onRef: (form) => {
              opts.form = form;
            },
            valueFormatter: (result) => {
              return result;
            },
            items: [
              {
                col: 0,
                values: [
                  {
                    type: 'select',
                    field: 'sheetLength',
                    name: I18n('EXCEL_SHEET_PER_ROW'),
                    map: Index.sheetLength,
                    binderType: 'number',
                  },
                  ...opts.filter,
                ],
              },
            ],
            operation: [
              {
                type: 'submit',
                label: I18n('EXCEL_EXPORT_START'),
              },
            ],
            onSubmit: (value) => {
              modal.destroy();
              let sheetLength = parseInt(value.sheetLength, 10);
              if (sheetLength <= 1 || isNaN(sheetLength)) {
                sheetLength = 1;
              }
              const loading = Modal.info({
                width: 700,
                title: I18n('WAIT_A_MOMENT'),
                className: 'vertical-center-modal',
                content: (
                  <Progress id={id} percent={0} status="active"/>
                ),
              });
              console.log(value);
              Index.pullLoop(opts.scope, value, {
                sheetLength: sheetLength,
                isZip: opts.isZip,
                sheet: opts.sheet,
              }, (percent) => {
                if (percent === -1) {
                  loading.destroy();
                  message.info(I18n('EXCEL_EXPORT_NO_DATA'));
                } else {
                  const apb = document.getElementById(id).getElementsByClassName('ant-progress-bg')[0];
                  const apt = document.getElementById(id).getElementsByClassName('ant-progress-text')[0];
                  apb.style.width = percent + '%';
                  apt.style.whiteSpace = 'nowrap';
                  apt.innerHTML = percent + '%';
                }
              });
            },
          }}
          />
        </div>
      ),
    });
  },

  push: (opts) => {
    function beforeUpload(info) {
      console.log(opts);
      Index.pushLoop(
        info,
        opts.set,
        opts.back,
      );
      setTimeout(() => {
        modal.destroy();
      }, 1000);
      return false;
    }

    const modal = Modal.info({
      width: 700,
      title: I18n('EXCEL_IMPORT'),
      maskClosable: true,
      className: 'vertical-center-modal hideFooter',
      content: (
        <div style={styles.excelBox}>
          <Alert
            message={I18n('EXCEL_IMPORT_DOWNLOAD_FILE_FOR_EXAMPLE')}
            description={(
              <div>
                <p>{opts.text}</p>
                {
                  opts.downloadText &&
                  <a style={{marginRight: '2rem'}} href={opts.downloadText} download={I18n('GUIDE_BOOK')}>
                    {I18n('EXCEL_DOWNLOAD_GUIDE_BOOK')}
                  </a>
                }
                {
                  opts.download &&
                  <a href={opts.download} download={I18n('GUIDE_TPL')}>
                    {I18n('EXCEL_DOWNLOAD_GUIDE_TPL')}
                  </a>
                }
              </div>
            )}
            type="info"
            showIcon
            banner
          />
          <Upload
            listType="text"
            action=""
            accept="application/vnd.ms-excel, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            beforeUpload={beforeUpload}
            multiple
          >
            <Button type="primary" style={{margin: "1rem 0 10px"}}>{I18n('UPLOAD_FILE')}</Button>
          </Upload>
        </div>
      ),
    });
  },

};

export default Index;
