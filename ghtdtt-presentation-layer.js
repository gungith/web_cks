    let danhSachCoQuanData = [];

    async function loadDanhSachCoQuan() {
      try {
        const data = await AppData.getOrganizations();
        if (Array.isArray(data)) {
          danhSachCoQuanData = data;
          const datalist = document.getElementById('danhSachCoQuan');
          datalist.innerHTML = '';
          const uniqueCoQuan = new Set();
          data.forEach(item => {
            const tenCQ = item.tenCoQuan || item.TenCoQuan || "";
            if (tenCQ) uniqueCoQuan.add(tenCQ);
          });
          uniqueCoQuan.forEach(tenCQ => {
            const option = document.createElement('option');
            option.value = tenCQ;
            datalist.appendChild(option);
          });
        }
      } catch (err) {
        console.error("Lỗi load danh sách cơ quan:", err);
      }
    }

    let templateArrayBuffer = null; 

    function clearSearch() {
      document.getElementById('searchSoCongVan').value = '';
      document.getElementById('btnClearSearch').classList.add('hidden');
      handleSearchInput();
    }

    async function loadSavedTemplate() {
      try {
        const template = await AppData.getTemplate();
        if (template) {
          templateArrayBuffer = template.data;
          const nameDisplay = document.getElementById('templateNameDisplay');
          nameDisplay.innerText = template.name;
          nameDisplay.classList.remove('text-indigo-600');
          nameDisplay.classList.add('text-emerald-600');
        }
        updateUploadButtonState();
      } catch (err) {
        console.error("Lỗi tải file mẫu:", err);
        showToast("Không thể tải file mẫu đã lưu!", "error");
      }
    }

    function loadTemplateFile(event) {
      const file = event.target.files[0];

      if (!file.name.endsWith('.docx') && !file.name.endsWith('.doc')) {
        showToast("Vui lòng chọn đúng file Word (.docx hoặc .doc)!", "error");
        return;
      }

      const reader = new FileReader();
      reader.onload = async function(e) {
        const data = e.target.result;
        try {
          await AppData.saveTemplate(file.name, data);
          templateArrayBuffer = data;
          const nameDisplay = document.getElementById('templateNameDisplay');
          nameDisplay.innerText = file.name;
          nameDisplay.classList.remove('text-indigo-600');
          nameDisplay.classList.add('text-emerald-600');
          updateUploadButtonState();
          showToast("Đã lưu file mẫu vào trình duyệt!", "success");
        } catch (err) {
          console.error("Lỗi lưu file mẫu:", err);
          showToast("Không thể lưu file mẫu!", "error");
        }
      };
      reader.onerror = () => showToast("Không thể đọc file mẫu!", "error");
      reader.readAsArrayBuffer(file);
    }

    async function clearTemplate() {
      if (!confirm("B?n c? ch?c ch?n mu?n x?a file m?u ?? l?u kh?ng?")) return;

      try {
        await AppData.deleteTemplate();
        templateArrayBuffer = null;
        document.getElementById('templateNameDisplay').innerText = "Chưa chọn file (Vui lòng chọn file)";
        document.getElementById('templateNameDisplay').classList.remove('text-emerald-600');
        document.getElementById('templateNameDisplay').classList.add('text-indigo-600');
        updateUploadButtonState();
        showToast("Đã xóa file mẫu!", "success");
      } catch (err) {
        console.error("Lỗi xóa file mẫu:", err);
        showToast("Không thể xóa file mẫu!", "error");
      }
    }

    function updateUploadButtonState() {
      const btn = document.getElementById('btnUploadTemplate');
      if (btn) {
        if (templateArrayBuffer) {
          btn.classList.add('hidden');
        } else {
          btn.classList.remove('hidden');
        }
      }
    }

    let lineChartInstance = null;
    let pieChartInstance = null;
    let completionChartInstance = null;
    let currentData = [];
    let recordActionCapabilities = { fullRecordEdit: false, deleteRow: false };
    let recordActionCapabilitiesLoaded = false;
    
    let currentPage = 1;
    const rowsPerPage = 10;

    const todayStr = new Date().toISOString().substring(0, 10);
    document.getElementById('inputNgay').value = todayStr;
    document.getElementById('inputNgayCongVan').value = todayStr;
    document.getElementById('filterCustomDate').value = todayStr;

    const datePickers = {};
    let editingRecordIndex = null;

    function openDatePicker(inputId) {
      if (datePickers[inputId]) datePickers[inputId].open();
    }
    
    function setupVietnameseFormValidation() {
      const forms = ['dataForm', 'editRecordForm']
        .map(formId => document.getElementById(formId))
        .filter(Boolean);

      forms.forEach(form => {
        form.addEventListener('invalid', event => {
          const field = event.target;
          if (!field.validity || typeof field.setCustomValidity !== 'function') return;

          if (field.validity.valueMissing) {
            field.setCustomValidity('Vui lòng điền thông tin này.');
          } else if (field.validity.typeMismatch && field.type === 'email') {
            field.setCustomValidity('Vui lòng nhập địa chỉ email hợp lệ.');
          } else if (field.validity.rangeUnderflow) {
            field.setCustomValidity(`Giá trị không được nhỏ hơn ${field.min}.`);
          } else if (field.validity.badInput) {
            field.setCustomValidity('Vui lòng nhập giá trị hợp lệ.');
          } else {
            field.setCustomValidity('Vui lòng kiểm tra lại thông tin đã nhập.');
          }
        }, true);

        const clearValidationMessage = event => {
          const field = event.target;
          if (typeof field.setCustomValidity === 'function') field.setCustomValidity('');
        };
        form.addEventListener('input', clearValidationMessage);
        form.addEventListener('change', clearValidationMessage);
      });
    }

    window.addEventListener('DOMContentLoaded', () => {
      if (typeof flatpickr === 'function') {
        const datePickerOptions = {
          altInput: true,
          altFormat: 'd/m/Y',
          dateFormat: 'Y-m-d',
          allowInput: true,
          disableMobile: true
        };
        ['inputNgay', 'inputNgayCongVan', 'inputNgayTraKQ'].forEach(inputId => {
          datePickers[inputId] = flatpickr(document.getElementById(inputId), datePickerOptions);
        });
        ['filterCustomDate'].forEach(inputId => {
          datePickers[inputId] = flatpickr(document.getElementById(inputId), {
            ...datePickerOptions,
            onChange: () => filterAndRender()
          });
        });
        datePickers.filterRangeDate = flatpickr(document.getElementById('filterRangeDate'), {
          ...datePickerOptions,
          mode: 'range',
          locale: flatpickr.l10ns.vn,
          onChange: (selectedDates, _dateString, instance) => {
            document.getElementById('filterStartDate').value = selectedDates[0]
              ? instance.formatDate(selectedDates[0], 'Y-m-d')
              : '';
            document.getElementById('filterEndDate').value = selectedDates[1]
              ? instance.formatDate(selectedDates[1], 'Y-m-d')
              : '';
            filterAndRender();
          }
        });
        ['editNgay', 'editNgayCongVan', 'editNgayTraKQ'].forEach(inputId => {
          datePickers[inputId] = flatpickr(document.getElementById(inputId), datePickerOptions);
        });
      } else {
        console.error('Không tải được Flatpickr; không thể khởi tạo trường ngày theo định dạng dd/mm/yyyy.');
      }

      setupVietnameseFormValidation();
      loadRecordActionCapabilities();
      loadData();
      loadDanhSachCoQuan();
      loadSavedTemplate();
    });

    function showToast(message, type = 'success') {
      const toast = document.getElementById('toast');
      const toastMessage = document.getElementById('toastMessage');
      const toastIcon = document.getElementById('toastIcon');

      toastMessage.innerText = message;

      if (type === 'success') {
        toast.className = "fixed bottom-5 right-5 transform translate-y-0 opacity-100 transition-all duration-300 z-50 px-4 py-3 rounded-xl shadow-lg border text-sm font-medium flex items-center gap-2 bg-emerald-50 border-emerald-200 text-emerald-800";
        toastIcon.innerText = "✅";
      } else {
        toast.className = "fixed bottom-5 right-5 transform translate-y-0 opacity-100 transition-all duration-300 z-50 px-4 py-3 rounded-xl shadow-lg border text-sm font-medium flex items-center gap-2 bg-red-50 border-red-200 text-red-800";
        toastIcon.innerText = "❌";
      }

      setTimeout(() => {
        toast.className = "fixed bottom-5 right-5 transform translate-y-20 opacity-0 transition-all duration-300 z-50 px-4 py-3 rounded-xl shadow-lg border text-sm font-medium flex items-center gap-2";
      }, 3000);
    }

    function formatReportISODate(date) {
      const year = date.getFullYear();
      const month = String(date.getMonth() + 1).padStart(2, '0');
      const day = String(date.getDate()).padStart(2, '0');
      return `${year}-${month}-${day}`;
    }

    function getReportDateRange(filters, records) {
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      let startDate = '';
      let endDate = '';

      if (filters.filterType === 'today') {
        startDate = endDate = formatReportISODate(today);
      } else if (filters.filterType === 'week') {
        const dayOfWeek = today.getDay() === 0 ? 6 : today.getDay() - 1;
        const start = new Date(today);
        start.setDate(today.getDate() - dayOfWeek);
        const end = new Date(start);
        end.setDate(start.getDate() + 6);
        startDate = formatReportISODate(start);
        endDate = formatReportISODate(end);
      } else if (filters.filterType === 'month') {
        startDate = formatReportISODate(new Date(today.getFullYear(), today.getMonth(), 1));
        endDate = formatReportISODate(new Date(today.getFullYear(), today.getMonth() + 1, 0));
      } else if (filters.filterType === 'custom') {
        startDate = endDate = filters.customDate || '';
      } else if (filters.filterType === 'range') {
        startDate = filters.startDate || '';
        endDate = filters.endDate || '';
      }

      if (!startDate || !endDate) {
        const recordDates = records
          .map(record => String(record.ngay || '').slice(0, 10))
          .filter(date => /^\d{4}-\d{2}-\d{2}$/.test(date))
          .sort();
        if (!startDate) startDate = recordDates[0] || '';
        if (!endDate) endDate = recordDates[recordDates.length - 1] || '';
      }

      return { startDate, endDate };
    }

    function openReportModal() {
      const filters = getCurrentFilters();
      const filteredRecords = AppBusiness.filterRecords(currentData, filters);
      const totals = AppBusiness.calculateKpis(filteredRecords);
      const totalRequests = totals.soToChuc + totals.soCaNhan;
      const dateRange = getReportDateRange(filters, filteredRecords);
      const fromDate = formatDateVN(dateRange.startDate) || '--';
      const toDate = formatDateVN(dateRange.endDate) || '--';
      const fromLabel = dateRange.startDate && dateRange.endDate
        && dateRange.startDate.slice(0, 4) === dateRange.endDate.slice(0, 4)
        ? fromDate.slice(0, 5)
        : fromDate;

      document.getElementById('reportContent').innerHTML = `
        <p>• Báo cáo anh từ <b>${fromLabel} đến ${toDate}</b>. Tổng đã tiếp nhận <b>${totalRequests.toLocaleString('vi-VN')}</b> yêu cầu gia hạn, thay đổi thông tin, trong đó bao gồm: <b>${totals.soCaNhan.toLocaleString('vi-VN')}</b> cá nhân, <b>${totals.soToChuc.toLocaleString('vi-VN')}</b> tổ chức.</p>
      `;

      const modal = document.getElementById('reportModal');
      modal.classList.remove('hidden');
      modal.classList.add('flex');
    }

    function closeReportModal() {
      const modal = document.getElementById('reportModal');
      modal.classList.remove('flex');
      modal.classList.add('hidden');
    }

    function hasUnsavedRowChanges(row, values) {
      return (Number(row.tongTiepNhan) || 0) !== values.tongTiepNhan
        || (Number(row.tongSanXuat) || 0) !== values.tongSanXuat
        || (Number(row.hoanThanh) || 0) !== values.hoanThanh
        || isRecordFlagEnabled(row.daLamVBPhucDap) !== values.daLamVBPhucDap
        || isRecordFlagEnabled(row.nopOnline) !== values.nopOnline;
    }

    function validateCompletionCount(index, row, action) {
      const completedInput = document.getElementById(`row_ht_${index}`);
      const completedCount = Number(completedInput.value) || 0;
      const requiredCount = (Number(row.soToChuc) || 0) + (Number(row.soCaNhan) || 0);
      if (completedCount === requiredCount) return true;

      alert(`KHÔNG THỂ ${action}!\n\nVui lòng kiểm tra lại số lượng hoàn thành!\n\nSau đó bấm "Cập nhật" để lưu hồ sơ!`);
      completedInput.classList.add('border-red-500', 'ring-2', 'ring-red-300');
      if (!completedInput.disabled) completedInput.focus();
      setTimeout(() => completedInput.classList.remove('border-red-500', 'ring-2', 'ring-red-300'), 2000);
      return false;
    }

    async function generateWordRow(index) {
      const row = currentData[index];
      if (!row) return;

      const rowInputs = {
        tongTiepNhan: document.getElementById(`row_tn_${index}`),
        tongSanXuat: document.getElementById(`row_sx_${index}`),
        hoanThanh: document.getElementById(`row_ht_${index}`)
      };
      if (Object.values(rowInputs).some(input => !input)) {
        showToast('Không thể kiểm tra trạng thái cập nhật của hồ sơ.', 'error');
        return;
      }

      const daLamVBPhucDapInput = document.getElementById(`row_pd_${index}`);
      const nopOnlineInput = document.getElementById(`row_online_${index}`);
      const visibleValues = {
        tongTiepNhan: Number(rowInputs.tongTiepNhan.value) || 0,
        tongSanXuat: Number(rowInputs.tongSanXuat.value) || 0,
        hoanThanh: Number(rowInputs.hoanThanh.value) || 0,
        daLamVBPhucDap: daLamVBPhucDapInput
          ? daLamVBPhucDapInput.checked
          : isRecordFlagEnabled(row.daLamVBPhucDap),
        nopOnline: nopOnlineInput
          ? nopOnlineInput.checked
          : isRecordFlagEnabled(row.nopOnline)
      };
      if (hasUnsavedRowChanges(row, visibleValues)) {
        showToast('Vui lòng bấm "Cập nhật" và lưu hồ sơ trước khi soạn văn bản.', 'error');
        return;
      }

      if (!validateCompletionCount(index, row, 'SOẠN VĂN BẢN')) return;

      if (!templateArrayBuffer) {
        showToast("Vui lòng chọn file mẫu .docx hoặc .doc ở góc trên màn hình trước!", "error");
        return;
      }

      if (!confirm("Công văn đã xử lý xong. Bạn có chắc muốn soạn Công văn phúc đáp không?")) {
        return;
      }

      try {
        showToast("Đang tạo file Word...", "success");

        const zip = new PizZip(templateArrayBuffer);
        const doc = new window.docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

        let ngayCongVanWord = "";
        if (row.ngayCongVan && row.ngayCongVan.includes('-')) {
          const parts = row.ngayCongVan.split('-');
          ngayCongVanWord = `${parts[2]} tháng ${parts[1]} năm ${parts[0]}`;
        } else {
          ngayCongVanWord = row.ngayCongVan || "";
        }

        let fullCoQuan = (row.coQuan || "").trim().replace(/\s+/g, ' ');

        const ubndRegex = /(Ủy|Uỷ)\s+ban\s+nhân\s+dân\s*|UBND\s*/gi;
        const duplicateUbndRegex = /((?:(?:Ủy|Uỷ)\s+ban\s+nhân\s+dân)|UBND)(?:\s+(?:(?:Ủy|Uỷ)\s+ban\s+nhân\s+dân)|UBND)+/gi;
        const hasDuplicateUbnd = /(?:(?:Ủy|Uỷ)\s+ban\s+nhân\s+dân|UBND)(?:\s+(?:Ủy|Uỷ)\s+ban\s+nhân\s+dân|UBND)+/i.test(fullCoQuan);
        const ubndMatch = fullCoQuan.match(/(Ủy|Uỷ)\s+ban\s+nhân\s+dân\s*|UBND\s*/i);
        const hasNameBeforeUbnd = ubndMatch
          && fullCoQuan.slice(0, ubndMatch.index).replace(/[\s\-–—,;:]+/g, '').trim();
        
        let coQuanWord = hasDuplicateUbnd
          ? fullCoQuan.replace(duplicateUbndRegex, '$1')
          : (hasNameBeforeUbnd ? fullCoQuan.replace(ubndRegex, '') : fullCoQuan);
        coQuanWord = coQuanWord
                                   .replace(/\s+/g, ' ')
                                   .replace(/\s*-\s*$/, '')
                                   .replace(/^\s*-\s*/, '')
                                   .trim();

        coQuanWord = coQuanWord.replace(/-\s*(phường|xã|quận|huyện|tỉnh|thành phố|thị xã|thị trấn)\b/gi, '$1');

        if (!coQuanWord) {
            coQuanWord = fullCoQuan;
        }

        const dataToFill = {
          soCongVan: row.soCongVan || "",
          ngayCongVan: ngayCongVanWord,
          coQuan: coQuanWord,
          trichYeu: row.trichYeu || "",
          soLuong: row.soLuong || 0
        };

        doc.render(dataToFill);
        const out = doc.getZip().generate({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });

        const sCV = removeVietnameseTones(row.soCongVan) || 'Moi';
        const cq = removeVietnameseTones(coQuanWord) || 'CoQuan';
        const fileName = `GHTĐTT_${sCV}_${cq}.docx`;
        saveAs(out, fileName);

        if (!isRecordFlagEnabled(row.daLamVBPhucDap)) {
          try {
            await AppData.updateRecord({
              action: "updateRow",
              rowId: row.rowId || '',
              soCongVan: row.soCongVan || '',
              tongTiepNhan: Number(row.tongTiepNhan) || 0,
              tongSanXuat: Number(row.tongSanXuat) || 0,
              hoanThanh: Number(row.hoanThanh) || 0,
              daLamVBPhucDap: true,
              nopOnline: isRecordFlagEnabled(row.nopOnline),
              mailDauMoi: row.mailDauMoi || '',
              ghiChu: row.ghiChu || ''
            });
            row.daLamVBPhucDap = true;
            filterAndRender();
            showToast("Đã soạn văn bản và đánh dấu VB Phúc đáp.", "success");
          } catch (err) {
            console.error("Lỗi lưu trạng thái VB Phúc đáp:", err);
            showToast("Đã tạo file nhưng không thể lưu trạng thái VB Phúc đáp.", "error");
          }
        } else {
          showToast("Đã tạo file Word.", "success");
        }
      } catch (err) {
        console.error(err);
        showToast("Lỗi khi tạo file!", "error");
      }
    }
    
    function removeVietnameseTones(str) {
      if (!str) return "";
      str = str.replace(/à|á|ạ|ả|ã|â|ầ|ấ|ậ|ẩ|ẫ|ă|ằ|ắ|ặ|ẳ|ẵ/g,"a"); 
      str = str.replace(/è|é|ẹ|ẻ|ẽ|ê|ề|ế|ệ|ể|ễ/g,"e"); 
      str = str.replace(/ì|í|ị|ỉ|ĩ/g,"i"); 
      str = str.replace(/ò|ó|ọ|ỏ|õ|ô|ồ|ố|ộ|ổ|ỗ|ơ|ờ|ớ|ợ|ở|ỡ/g,"o"); 
      str = str.replace(/ù|ú|ụ|ủ|ũ|ư|ừ|ứ|ự|ử|ữ/g,"u"); 
      str = str.replace(/ỳ|ý|ỵ|ỷ|ỹ/g,"y"); 
      str = str.replace(/đ/g,"d");
      str = str.replace(/À|Á|Ạ|Ả|Ã|Â|Ầ|Ấ|Ậ|Ẩ|Ẫ|Ă|Ằ|Ắ|Ặ|Ẳ|Ẵ/g, "A");
      str = str.replace(/È|É|Ẹ|Ẻ|Ẽ|Ê|Ề|Ế|Ệ|Ể|Ễ/g, "E");
      str = str.replace(/Ì|Í|Ị|Ỉ|Ĩ/g, "I");
      str = str.replace(/Ò|Ó|Ọ|Ỏ|Õ|Ô|Ồ|Ố|Ộ|Ổ|Ỗ|Ơ|Ờ|Ớ|Ợ|Ở|Ỡ/g, "O");
      str = str.replace(/Ù|Ú|Ụ|Ủ|Ũ|Ư|Ừ|Ứ|Ự|Ử|Ữ/g, "U");
      str = str.replace(/Ỳ|Ý|Ỵ|Ỷ|Ỹ/g, "Y");
      str = str.replace(/Đ/g, "D");
      str = str.replace(/[^a-zA-Z0-9 ]/g, ""); 
      str = str.replace(/\s+/g, "_"); 
      return str;
    }
	
    async function loadData(showSuccessToast = true) {
      try {
        currentData = await AppData.getRecords();
        filterAndRender();
        if (showSuccessToast) {
          showToast("Cập nhật dữ liệu thành công!", "success");
        }
      } catch (err) {
        console.error("Lỗi tải dữ liệu:", err);
        showToast("Không thể kết nối để tải dữ liệu!", "error");
      }
    }

    async function loadRecordActionCapabilities() {
      try {
        const capabilities = await AppData.getRecordActionCapabilities();
        recordActionCapabilities = {
          fullRecordEdit: capabilities && capabilities.fullRecordEdit === true,
          deleteRow: capabilities && capabilities.deleteRow === true
        };
      } catch (err) {
        console.error('Không thể kiểm tra hỗ trợ sửa/xóa công văn:', err);
      } finally {
        recordActionCapabilitiesLoaded = true;
      }
    }

    function handleFilterChange() {
      const filterType = document.getElementById('filterType').value;
      const customContainer = document.getElementById('customDateContainer');
      const rangeContainer = document.getElementById('rangeDateContainer');
      
      customContainer.classList.add('hidden');
      rangeContainer.classList.add('hidden');

      if (filterType === 'custom') {
        customContainer.classList.remove('hidden');
      } else if (filterType === 'range') {
        rangeContainer.classList.remove('hidden');
      }
      
      currentPage = 1;
      filterAndRender();
    }

    function handleSearchInput() {
      const searchInput = document.getElementById('searchSoCongVan');
      const btnClear = document.getElementById('btnClearSearch');
      
      if (searchInput.value.length > 0) {
        btnClear.classList.remove('hidden');
      } else {
        btnClear.classList.add('hidden');
      }

      currentPage = 1;
      filterAndRender();
    }

    function getCurrentFilters() {
      return {
        filterType: document.getElementById('filterType').value,
        searchQuery: document.getElementById('searchSoCongVan').value.toLowerCase().trim(),
        customDate: document.getElementById('filterCustomDate').value,
        startDate: document.getElementById('filterStartDate').value,
        endDate: document.getElementById('filterEndDate').value,
        overdueOnly: isOverdueFilterActive,
        unrepliedOnly: isUnrepliedFilterActive
      };
    }

    function filterAndRender() {
      const filters = getCurrentFilters();
      const filteredData = AppBusiness.filterRecords(currentData, filters);
      const summaryData = filters.searchQuery
        ? AppBusiness.filterRecords(currentData, { ...filters, searchQuery: '' })
        : filteredData;
      updateKPIs(summaryData);
      renderTableWithPagination(filteredData);
      renderCharts(summaryData);
      checkOverdueNotifications();
    }

    function updateKPIs(data) {
      const totals = AppBusiness.calculateKpis(data);
      const values = {
        kpiTotalDocuments: totals.soVanBan,
        kpiTotalDevice: totals.soToChuc + totals.soCaNhan,
        kpiTotalTN: totals.tongTiepNhan,
        kpiTotalSX: totals.tongSanXuat,
        kpiTotalCaNhan: totals.soCaNhan,
        kpiTotalToChuc: totals.soToChuc,
        kpiTotalHoanThanh: totals.soCongVanHoanThanh,
        kpiTotalDeviceHoanThanh: totals.hoanThanh
      };
      Object.entries(values).forEach(([id, value]) => {
        document.getElementById(id).innerText = value.toLocaleString('vi-VN');
      });

      const totalDevices = totals.soToChuc + totals.soCaNhan;
      const completedDevices = totals.hoanThanh;
      const completionRate = totalDevices > 0
        ? Math.min(100, Math.round((completedDevices / totalDevices) * 100))
        : 0;
      document.getElementById('deviceCompletionSummary').innerText =
        `${completedDevices.toLocaleString('vi-VN')} / ${totalDevices.toLocaleString('vi-VN')} thiết bị hoàn thành`;
      document.getElementById('deviceCompletionRate').innerText = `${completionRate}%`;
      document.getElementById('deviceCompletionBar').style.width = `${completionRate}%`;
      document.getElementById('deviceCompletionBar').parentElement.setAttribute('aria-valuenow', String(completionRate));
    }

    function renderTableWithPagination(data) {
      const tbody = document.getElementById('tableBody');
      tbody.innerHTML = "";

      if (!data || data.length === 0) {
        tbody.innerHTML = '<tr><td colspan="17" class="text-center py-4 text-gray-400">Không tìm thấy công văn phù hợp</td></tr>';
        document.getElementById('paginationInfo').innerText = "Hiển thị 0 - 0 của 0 công văn";
        document.getElementById('paginationControls').innerHTML = "";
        return;
      }

      const reversedData = data.slice().reverse();
      const totalPages = Math.ceil(reversedData.length / rowsPerPage);
      if (currentPage > totalPages) currentPage = totalPages;
      if (currentPage < 1) currentPage = 1;

      const startIndex = (currentPage - 1) * rowsPerPage;
      const endIndex = Math.min(startIndex + rowsPerPage, reversedData.length);
      const paginatedData = reversedData.slice(startIndex, endIndex);

      document.getElementById('paginationInfo').innerText = `Hiển thị ${startIndex + 1} - ${endIndex} của ${reversedData.length} công văn`;

      const today = new Date();
      today.setHours(0, 0, 0, 0);

      paginatedData.forEach((row, idx) => {
        const actualIndex = currentData.indexOf(row);
        const rowId = row.rowId || '';

        const ngayTraKQVal = formatDateVN(row.ngayTraKetQua) || '-';
        const soCVVal = row.soCongVan || '';
        const ngayCVVal = formatDateVN(row.ngayCongVan) || '-';
        const coQuanVal = row.coQuan || '';
        
        const soTCVal = row.soToChuc || 0;
        const soCNVal = row.soCaNhan || 0;
        
        const tongTiepNhan = Number(row.tongTiepNhan) || 0;
        const tongSanXuat = Number(row.tongSanXuat) || 0;
        
        const htVal = row.hoanThanh || 0;

        const isXL = row.daXuLy;
        const isPD = (row.daLamVBPhucDap === true || row.daLamVBPhucDap === 1 || row.daLamVBPhucDap === "1" || row.daLamVBPhucDap === "true" || row.daLamVBPhucDap === "Có");
        const isNopOnline = (row.nopOnline === true || row.nopOnline === 1 || row.nopOnline === "1" || row.nopOnline === "true" || row.nopOnline === "Có");
        const isDone = isPD || isNopOnline;
  
        const mailDauMoi = row.mailDauMoi || '';
        const ghiChu = row.ghiChu || '';
        const trichYeuVal = row.trichYeu || '';
        
        let isOverdue = false;
        if (row.ngayTraKetQua && !isPD && !isNopOnline) {
          const deadline = new Date(row.ngayTraKetQua);
          deadline.setHours(0, 0, 0, 0);
          if (deadline < today) { 
            isOverdue = true;
          }
        }

        const rowClass = isOverdue ? "bg-red-50 hover:bg-red-100 border-b border-red-200 text-red-700" : "hover:bg-gray-50 border-b border-gray-100";

        const safeSoCV = String(soCVVal).replace(/'/g, "\\'");

        const disabledAttr = isDone ? 'disabled' : '';
        const inputDisabledClass = isDone ? 'bg-gray-100 text-gray-500 cursor-not-allowed' : '';

        const actionButton = `
            <div class="flex flex-col gap-1 items-center action-buttons bg-white p-1 rounded shadow-sm border border-gray-100"> 
              ${isDone 
                ? '<span class="text-[10px] font-semibold text-emerald-600 bg-emerald-50 px-1 py-0.5 rounded border border-emerald-200 text-center w-full block">Đã xong</span>'
                : `<button onclick="updateRowData(${actualIndex}, ${rowId}, '${safeSoCV}')" class="px-2 py-1 bg-indigo-600 text-white rounded hover:bg-indigo-700 text-xs font-medium transition shadow-sm w-full whitespace-nowrap">Cập nhật</button>`
              }
            ${isDone ? '' : `<button onclick="generateWordRow(${actualIndex})" class="px-2 py-1 bg-teal-600 text-white rounded hover:bg-teal-700 text-xs font-medium transition shadow-sm w-full flex items-center justify-center gap-1 whitespace-nowrap">
              <span>📄 Soạn VB</span>
            </button>`}
            ${isDone ? '' : `
              <button type="button" onclick="openEditRecordModal(${actualIndex})" title="Chỉnh sửa toàn bộ công văn" class="px-2 py-1 bg-amber-500 text-white rounded hover:bg-amber-600 text-xs font-medium transition shadow-sm w-full whitespace-nowrap">
                ✏️ Sửa
              </button>
              <button type="button" onclick="deleteRecordRow(${actualIndex})" title="Xóa công văn" class="px-2 py-1 bg-rose-600 text-white rounded hover:bg-rose-700 text-xs font-medium transition shadow-sm w-full whitespace-nowrap">
                🗑️ Xóa
              </button>
            `}
          </div>
        `;

        const tr = document.createElement('tr');
        tr.className = rowClass; 
               tr.innerHTML = `
          <td class="py-3 px-2 whitespace-nowrap font-medium">${ngayTraKQVal}</td>
          <td class="py-3 px-2 font-bold whitespace-nowrap">${soCVVal}</td>
          <td class="py-3 px-2 whitespace-nowrap">${ngayCVVal}</td>
          <td class="py-3 px-2 text-xs font-semibold">${coQuanVal}</td>
          <td class="py-3 px-2 text-center font-semibold">${Number(soTCVal).toLocaleString('vi-VN')}</td>
          <td class="py-3 px-2 text-center font-semibold">${Number(soCNVal).toLocaleString('vi-VN')}</td>
          
          <td class="py-2 px-1 text-center">
            <input type="number" id="row_tn_${actualIndex}" value="${tongTiepNhan}" min="0" ${disabledAttr} class="table-number-input table-number-input--received w-14 px-1 py-1 text-center border border-sky-300 rounded focus:outline-none focus:ring-1 focus:ring-sky-500 font-bold text-sky-700 text-xs ${inputDisabledClass}">
          </td>
          <td class="py-2 px-1 text-center">
            <input type="number" id="row_sx_${actualIndex}" value="${tongSanXuat}" min="0" ${disabledAttr} class="table-number-input table-number-input--production w-14 px-1 py-1 text-center border border-amber-300 rounded focus:outline-none focus:ring-1 focus:ring-amber-500 font-bold text-amber-700 text-xs ${inputDisabledClass}">
          </td>
          
          <td class="py-2 px-1 text-center">
            <div class="table-number-wrap relative w-16 mx-auto flex items-center">
              <input type="number" id="row_ht_${actualIndex}" value="${htVal}" min="0" ${disabledAttr} class="table-number-input table-number-input--completed w-full pl-1 pr-5 py-1 text-center border border-emerald-300 rounded focus:outline-none focus:ring-1 focus:ring-emerald-500 font-bold text-emerald-700 text-xs ${inputDisabledClass}">
               ${(isDone || disabledAttr) ? '' : `<button onclick="autoFillMaxRow(${actualIndex}, '${safeSoCV}')" title="Lấy (Tổ chức + Cá nhân) điền vào tất cả" class="table-number-autofill absolute right-0.5 top-1/2 transform -translate-y-1/2 px-0.5 bg-emerald-500 text-white rounded hover:bg-emerald-600 transition text-[9px] shadow-sm">⚡</button>`}
            </div>
          </td>

          <td class="py-3 px-2 text-center">${isXL ? '<span class="text-indigo-600 font-bold">✓</span>' : '<span class="text-gray-300">-</span>'}</td>
          
          <td class="py-2 px-1 text-center">
            <div class="min-w-[140px] max-w-[220px] whitespace-normal break-words text-xs" title="${mailDauMoi}">${mailDauMoi || '-'}</div>
          </td>

          <td class="py-2 px-1 text-center">
             <div class="min-w-[120px] max-w-[220px] whitespace-normal break-words text-xs" title="${ghiChu}">${ghiChu || '-'}</div>
          </td>

          <td class="py-3 px-2 text-center">
            ${isDone 
              ? (isPD ? '<span class="text-emerald-600 font-bold text-lg">✓</span>' : '<span class="text-gray-300">-</span>') 
              : `<input type="checkbox" id="row_pd_${actualIndex}" ${isPD ? 'checked' : ''} 
                  onclick="handleCheckboxClick(this, ${actualIndex}, '${safeSoCV}')" 
                  class="w-4 h-4 text-blue-600 rounded border-gray-300 focus:ring-blue-500 cursor-pointer">`
            }
          </td>
          <td class="py-3 px-2 text-center">
            ${isDone 
              ? (isNopOnline ? '<span class="text-emerald-600 font-bold text-lg">✓</span>' : '<span class="text-gray-300">-</span>') 
              : `<input type="checkbox" id="row_online_${actualIndex}" ${isNopOnline ? 'checked' : ''} 
                  onclick="handleNopOnlineClick(this, ${actualIndex}, '${safeSoCV}')" 
                  class="w-4 h-4 text-blue-600 rounded border-gray-300 focus:ring-blue-500 cursor-pointer">`
            }
          </td>

          <td class="py-3 px-2">
            <div class="min-w-[220px] max-w-[420px] whitespace-normal break-words" title="${trichYeuVal}">${trichYeuVal}</div>
          </td>
          <td class="py-2 px-2 text-center">
            ${actionButton}
          </td>
        `;
        tbody.appendChild(tr);
      });

      renderPaginationControls(totalPages);
      syncScrollbars();
    }
  
    function formatDateVN(dateString) {
      if (!dateString || dateString === '-') return '-';
      if (dateString.includes('/')) return dateString;
      
      const parts = dateString.split('-');
      if (parts.length === 3) {
        return `${parts[2]}/${parts[1]}/${parts[0]}`;
      }
      return dateString;
    }

    function renderPaginationControls(totalPages) {
      const controlsContainer = document.getElementById('paginationControls');
      controlsContainer.innerHTML = "";

      if (totalPages <= 1) return;

      const prevBtn = document.createElement('button');
      prevBtn.className = `px-3 py-1 rounded text-xs font-medium border ${currentPage === 1 ? 'bg-gray-50 text-gray-300 border-gray-200 cursor-not-allowed' : 'bg-white text-gray-700 border-gray-300 hover:bg-gray-50'}`;
      prevBtn.innerText = "‹ Trước";
      prevBtn.disabled = currentPage === 1;
      prevBtn.onclick = () => {
        if (currentPage > 1) {
          currentPage--;
          filterAndRender();
        }
      };
      controlsContainer.appendChild(prevBtn);

      for (let i = 1; i <= totalPages; i++) {
        if (i === 1 || i === totalPages || (i >= currentPage - 1 && i <= currentPage + 1)) {
          const pageBtn = document.createElement('button');
          pageBtn.className = `px-3 py-1 rounded text-xs font-medium border ${i === currentPage ? 'bg-indigo-600 text-white border-indigo-600' : 'bg-white text-gray-700 border-gray-300 hover:bg-gray-50'}`;
          pageBtn.innerText = i;
          pageBtn.onclick = () => {
            currentPage = i;
            filterAndRender();
          };
          controlsContainer.appendChild(pageBtn);
        } else if (i === currentPage - 2 || i === currentPage + 2) {
          const span = document.createElement('span');
          span.className = "px-2 text-gray-400 text-xs";
          span.innerText = "...";
          controlsContainer.appendChild(span);
        }
      }

      const nextBtn = document.createElement('button');
      nextBtn.className = `px-3 py-1 rounded text-xs font-medium border ${currentPage === totalPages ? 'bg-gray-50 text-gray-300 border-gray-200 cursor-not-allowed' : 'bg-white text-gray-700 border-gray-300 hover:bg-gray-50'}`;
      nextBtn.innerText = "Sau ›";
      nextBtn.disabled = currentPage === totalPages;
      nextBtn.onclick = () => {
        if (currentPage < totalPages) {
          currentPage++;
          filterAndRender();
        }
      };
      controlsContainer.appendChild(nextBtn);
    }

    function renderCharts(data) {
      const dateMap = {};
      const coQuanMap = {};

      data.forEach(item => {
        const date = item.ngay;
        const coQuan = item.coQuan ? item.coQuan.trim() : "Chưa phân loại";
        
        const soToChuc = Number(item.soToChuc) || 0;
        const soCaNhan = Number(item.soCaNhan) || 0;
        const tongSoLuong = soToChuc + soCaNhan;

        if (date) {
          dateMap[date] = (dateMap[date] || 0) + 1;
        }
        coQuanMap[coQuan] = (coQuanMap[coQuan] || 0) + tongSoLuong;
      });

      const sortedDates = Object.keys(dateMap).sort();
      const dateValues = sortedDates.map(d => dateMap[d]);

      const ctxLine = document.getElementById('lineChart').getContext('2d');
      if (lineChartInstance) lineChartInstance.destroy();
      lineChartInstance = new Chart(ctxLine, {
        type: 'line',
        data: {
          labels: sortedDates,
          datasets: [{
            label: 'Số lượng công văn',
            data: dateValues,
            borderColor: '#6366f1',
            backgroundColor: 'rgba(99, 102, 241, 0.1)',
            fill: true,
            tension: 0.3
          }]
        },
        options: { 
          responsive: true, 
          maintainAspectRatio: false,
          scales: {
            y: {
              beginAtZero: true,
              ticks: {
                stepSize: 1,
                precision: 0
              }
            }
          }
        }
      });

      const coQuanLabels = Object.keys(coQuanMap);
      const coQuanValues = coQuanLabels.map(cq => coQuanMap[cq]);
      const backgroundColors = ['#6366f1', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#06b6d4', '#ec4899', '#3b82f6', '#14b8a6', '#f43f5e', '#84cc16'];

      const ctxPie = document.getElementById('pieChart').getContext('2d');
      if (pieChartInstance) pieChartInstance.destroy();
      pieChartInstance = new Chart(ctxPie, {
        type: 'doughnut',
        data: {
          labels: coQuanLabels,
          datasets: [{
            data: coQuanValues,
            backgroundColor: backgroundColors
          }]
        },
        options: { 
          responsive: true, 
          maintainAspectRatio: false,
          plugins: {
            legend: { display: false },
            tooltip: {
              callbacks: {
                label: function(context) {
                  let label = context.label || '';
                  let value = context.raw || 0;
                  return ` ${label}: ${value.toLocaleString('vi-VN')} tổng số lượng`;
                }
              }
            }
          }
        }
      });

      const legendContainer = document.getElementById('pieLegendCustom');
      legendContainer.innerHTML = '';
      coQuanLabels.forEach((label, i) => {
        const color = backgroundColors[i % backgroundColors.length];
        const value = coQuanValues[i];
        
        const itemDiv = document.createElement('div');
        itemDiv.className = "flex items-center justify-between gap-2 py-1 px-1.5 hover:bg-gray-50 rounded cursor-pointer";
        itemDiv.innerHTML = `
          <div class="flex items-center gap-2 overflow-hidden">
            <span class="w-3 h-3 rounded-full flex-shrink-0" style="background-color: ${color};"></span>
            <span class="truncate text-gray-700" title="${label}">${label}</span>
          </div>
          <span class="font-semibold text-gray-900 flex-shrink-0">${value.toLocaleString('vi-VN')}</span>
        `;
        legendContainer.appendChild(itemDiv);
      });

      renderCompletionChart(data);
    }    

    function renderCompletionChart(data) {
      const completion = AppBusiness.calculateDocumentCompletion(data);
      document.getElementById('completionRate').innerText = completion.percentage + '%';
      document.getElementById('completionSummary').innerText =
        completion.completed.toLocaleString('vi-VN') + ' / ' + completion.total.toLocaleString('vi-VN') + ' công văn';

      const ctx = document.getElementById('completionChart').getContext('2d');
      if (completionChartInstance) completionChartInstance.destroy();
      completionChartInstance = new Chart(ctx, {
        type: 'doughnut',
        data: {
          labels: ['Đã hoàn thành', 'Chưa hoàn thành'],
          datasets: [{
            data: completion.total
              ? [completion.completed, completion.total - completion.completed]
              : [0, 1],
            backgroundColor: completion.total
              ? ['#10b981', '#d1d5db']
              : ['#d1d5db', '#d1d5db'],
            borderWidth: 0
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          cutout: '72%',
          plugins: {
            legend: { display: false },
            tooltip: {
              callbacks: {
                label: function(context) {
                  if (!completion.total) return ' Chưa có dữ liệu';
                  return ' ' + context.label + ': ' + Number(context.raw || 0).toLocaleString('vi-VN');
                }
              }
            }
          }
        }
      });
    }

    function normalizeSingleLineText(value) {
      return String(value || '')
        .replace(/[\r\n\u2028\u2029]+/g, ' ')
        .replace(/[ \t]+/g, ' ')
        .trim();
    }

    async function submitData(e) {
      e.preventDefault();
      if (!confirm("Bạn có chắc muốn thêm công văn này không?")) return;

      const btn = document.getElementById('btnSubmit');
      btn.disabled = true;
      btn.innerText = "Đang gửi...";

      const payload = AppBusiness.createNewRecord({
        ngay: document.getElementById('inputNgay').value,
        ngayCongVan: document.getElementById('inputNgayCongVan').value,
        ngayTraKetQua: document.getElementById('inputNgayTraKQ').value,
        soCongVan: document.getElementById('inputSoCongVan').value,
        coQuan: document.getElementById('inputCoQuan').value.trim(),
        donVi: document.getElementById('inputDonVi').value.trim(),
        trichYeu: normalizeSingleLineText(document.getElementById('inputTrichYeu').value),
        soToChuc: Number(document.getElementById('inputSoToChuc').value) || 0,
        soCaNhan: Number(document.getElementById('inputSoCaNhan').value) || 0,
        mailDauMoi: document.getElementById('inputMailDauMoi').value
      });

      try {
        await AppData.createRecord(payload);
        document.getElementById('inputSoCongVan').value = '';
        document.getElementById('inputDonVi').value = '';
        document.getElementById('inputCoQuan').value = '';
        document.getElementById('inputTrichYeu').value = '';
        resizeTrichYeuInput();
        toggleClearTrichYeuBtn();
        document.getElementById('inputSoToChuc').value = '';
        document.getElementById('inputSoCaNhan').value = '';
        if (datePickers.inputNgayTraKQ) {
          datePickers.inputNgayTraKQ.clear();
        } else {
          document.getElementById('inputNgayTraKQ').value = '';
        }
        document.getElementById('inputMailDauMoi').value = '';
        showToast("Thêm mới thành công!", "success");
        btn.disabled = false;
        btn.innerText = "Thêm mới công văn";
        void loadData(false);
      } catch (err) {
        console.error("Lỗi khi gửi dữ liệu:", err);
        showToast("Gửi dữ liệu thất bại, vui lòng thử lại!", "error");
        btn.disabled = false;
        btn.innerText = "Thêm mới công văn";
      }
    }

    async function updateRowData(index, rowId, soCongVan) {
      const row = currentData[index];
      if (!row) {
        showToast('Không tìm thấy công văn cần cập nhật.', 'error');
        return;
      }

      const tnTotal = document.getElementById(`row_tn_${index}`).value === "" ? 0 : Number(document.getElementById(`row_tn_${index}`).value);
      const sxTotal = document.getElementById(`row_sx_${index}`).value === "" ? 0 : Number(document.getElementById(`row_sx_${index}`).value);
      const ht = document.getElementById(`row_ht_${index}`).value === "" ? 0 : Number(document.getElementById(`row_ht_${index}`).value);
      if (!validateCompletionCount(index, row, 'CẬP NHẬT')) return;
      
      const daPhucDap = document.getElementById(`row_pd_${index}`).checked;
      const isNopOnline = document.getElementById(`row_online_${index}`).checked;

      const mailDauMoi = row.mailDauMoi || '';
      const ghiChu = row.ghiChu || '';

      const payload = {
        action: "updateRow",
        rowId: rowId,
        soCongVan: soCongVan,
        tongTiepNhan: tnTotal,   
        tongSanXuat: sxTotal,    
        hoanThanh: ht,
        daLamVBPhucDap: daPhucDap,
        nopOnline: isNopOnline,
        mailDauMoi: mailDauMoi,
        ghiChu: ghiChu
      };

      try {
        await AppData.updateRecord(payload);

        row.tongTiepNhan = tnTotal;
        row.tongSanXuat = sxTotal;
        row.hoanThanh = ht;
        row.daLamVBPhucDap = daPhucDap;
        row.nopOnline = isNopOnline;
        filterAndRender();
        showToast(`Đã cập nhật công văn ${soCongVan}!`, "success");
        setTimeout(async () => {
          await loadData();
        }, 1500);
      } catch (err) {
        console.error("Lỗi cập nhật:", err);
        showToast("Cập nhật thất bại!", "error");
      }
    }

    function normalizeEditDate(value) {
      const date = String(value || '').trim();
      const dmy = date.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (dmy) {
        return `${dmy[3]}-${dmy[2].padStart(2, '0')}-${dmy[1].padStart(2, '0')}`;
      }
      return date;
    }

    function isRecordFlagEnabled(value) {
      return value === true || value === 1 || value === '1' || String(value).toLowerCase() === 'true' || value === 'Có' || value === 'X' || value === 'x';
    }

    function setEditDate(inputId, value) {
      const picker = datePickers[inputId];
      const date = normalizeEditDate(value);
      if (!picker) return;
      if (date) {
        picker.setDate(date, false, 'Y-m-d');
      } else {
        picker.clear();
      }
    }

    function openEditRecordModal(index) {
      if (!recordActionCapabilitiesLoaded) {
        showToast('Đang kiểm tra chức năng trên máy chủ, vui lòng thử lại.', 'error');
        return;
      }
      if (!recordActionCapabilities.fullRecordEdit) {
        showToast('Cần cập nhật và triển khai Apps Script trước khi sửa toàn bộ công văn.', 'error');
        return;
      }

      const row = currentData[index];
      if (!row) {
        showToast('Không tìm thấy công văn cần chỉnh sửa.', 'error');
        return;
      }

      editingRecordIndex = index;
      setEditDate('editNgay', row.ngay);
      setEditDate('editNgayCongVan', row.ngayCongVan);
      setEditDate('editNgayTraKQ', row.ngayTraKetQua);
      document.getElementById('editSoCongVan').value = row.soCongVan || '';
      document.getElementById('editCoQuan').value = row.coQuan || '';
      document.getElementById('editSoToChuc').value = Number(row.soToChuc) || 0;
      document.getElementById('editSoCaNhan').value = Number(row.soCaNhan) || 0;
      document.getElementById('editTongTiepNhan').value = Number(row.tongTiepNhan) || 0;
      document.getElementById('editTongSanXuat').value = Number(row.tongSanXuat) || 0;
      document.getElementById('editHoanThanh').value = Number(row.hoanThanh) || 0;
      document.getElementById('editMailDauMoi').value = row.mailDauMoi || '';
      document.getElementById('editGhiChu').value = row.ghiChu || '';
      document.getElementById('editTrichYeu').value = row.trichYeu || '';
      document.getElementById('editDaXuLy').checked = isRecordFlagEnabled(row.daXuLy);
      document.getElementById('editPhucDap').checked = isRecordFlagEnabled(row.daLamVBPhucDap);
      document.getElementById('editNopOnline').checked = isRecordFlagEnabled(row.nopOnline);
      document.getElementById('editRecordTitle').textContent = `Chỉnh sửa công văn ${row.soCongVan || ''}`;
      document.getElementById('editRecordForm').querySelectorAll('input, textarea').forEach(field => field.setCustomValidity(''));

      const modal = document.getElementById('editRecordModal');
      modal.classList.remove('hidden');
      modal.classList.add('flex');
    }

    function closeEditRecordModal() {
      const modal = document.getElementById('editRecordModal');
      modal.classList.remove('flex');
      modal.classList.add('hidden');
      editingRecordIndex = null;
    }

    async function submitEditRecord(event) {
      event.preventDefault();
      if (!recordActionCapabilitiesLoaded) {
        showToast('Đang kiểm tra chức năng trên máy chủ, vui lòng thử lại.', 'error');
        return;
      }
      const row = currentData[editingRecordIndex];
      if (!row) {
        showToast('Không tìm thấy công văn cần chỉnh sửa.', 'error');
        closeEditRecordModal();
        return;
      }

      if (!recordActionCapabilities.fullRecordEdit) {
        showToast('Cần cập nhật và triển khai Apps Script trước khi lưu chỉnh sửa.', 'error');
        return;
      }

      const soToChuc = Number(document.getElementById('editSoToChuc').value) || 0;
      const soCaNhan = Number(document.getElementById('editSoCaNhan').value) || 0;
      const hoanThanh = Number(document.getElementById('editHoanThanh').value) || 0;
      const daLamVBPhucDap = document.getElementById('editPhucDap').checked;
      const nopOnline = document.getElementById('editNopOnline').checked;
      if ((daLamVBPhucDap || nopOnline) && hoanThanh !== soToChuc + soCaNhan) {
        showToast('Số hoàn thành phải bằng tổng tổ chức và cá nhân khi đánh dấu đã hoàn thành.', 'error');
        return;
      }

      const saveButton = document.getElementById('btnSaveEditRecord');
      saveButton.disabled = true;
      saveButton.innerText = 'Đang lưu...';

      const payload = {
        action: 'updateRow',
        fullEdit: true,
        rowId: row.rowId,
        id: row.id,
        ngay: document.getElementById('editNgay').value,
        soCongVan: document.getElementById('editSoCongVan').value.trim(),
        ngayCongVan: document.getElementById('editNgayCongVan').value,
        coQuan: document.getElementById('editCoQuan').value.trim(),
        soToChuc,
        soCaNhan,
        daXuLy: document.getElementById('editDaXuLy').checked,
        ngayTraKetQua: document.getElementById('editNgayTraKQ').value,
        tongTiepNhan: Number(document.getElementById('editTongTiepNhan').value) || 0,
        tongSanXuat: Number(document.getElementById('editTongSanXuat').value) || 0,
        hoanThanh,
        daLamVBPhucDap,
        nopOnline,
        mailDauMoi: document.getElementById('editMailDauMoi').value.trim(),
        ghiChu: document.getElementById('editGhiChu').value.trim(),
        trichYeu: normalizeSingleLineText(document.getElementById('editTrichYeu').value)
      };

      try {
        await AppData.updateRecord(payload);
        closeEditRecordModal();
        showToast(`Đã gửi yêu cầu cập nhật công văn ${payload.soCongVan}.`, 'success');
        setTimeout(() => loadData(), 1200);
      } catch (err) {
        console.error('Lỗi chỉnh sửa công văn:', err);
        showToast('Không thể gửi yêu cầu chỉnh sửa công văn.', 'error');
      } finally {
        saveButton.disabled = false;
        saveButton.innerText = 'Lưu thay đổi';
      }
    }

    async function deleteRecordRow(index) {
      if (!recordActionCapabilitiesLoaded) {
        showToast('Đang kiểm tra chức năng trên máy chủ, vui lòng thử lại.', 'error');
        return;
      }
      if (!recordActionCapabilities.deleteRow) {
        showToast('Cần cập nhật và triển khai Apps Script trước khi xóa công văn.', 'error');
        return;
      }

      const row = currentData[index];
      if (!row) {
        showToast('Không tìm thấy công văn cần xóa.', 'error');
        return;
      }

      const soCongVan = row.soCongVan || '(chưa có số)';
      if (!confirm(`Bạn có chắc chắn muốn xóa công văn ${soCongVan} không? Thao tác này không thể hoàn tác.`)) return;

      try {
        await AppData.deleteRecord({
          action: 'deleteRow',
          id: row.id,
          rowId: row.rowId,
          soCongVan: row.soCongVan
        });
        showToast(`Đã gửi yêu cầu xóa công văn ${soCongVan}.`, 'success');
        setTimeout(() => loadData(), 1200);
      } catch (err) {
        console.error('Lỗi xóa công văn:', err);
        showToast('Không thể gửi yêu cầu xóa công văn.', 'error');
      }
    }

    // ==========================================
    // LOGIC CHUÔNG THÔNG BÁO QUÁ HẠN
    // ==========================================
    
    function updateOverdueTicker(overdueList) {
      const ticker = document.getElementById('overdueTicker');
      const track = document.getElementById('overdueTickerTrack');
      const count = document.getElementById('overdueTickerCount');
      track.innerHTML = '';

      if (overdueList.length === 0) {
        ticker.classList.add('hidden');
        return;
      }

      count.textContent = overdueList.length;
      ticker.classList.remove('hidden');

      const fieldsFor = item => [
        ['Số công văn', item.soCongVan],
        ['Ngày công văn', formatDateVN(item.ngayCongVan)],
        ['Cơ quan', item.coQuan],
        ['Hạn trả kết quả', formatDateVN(item.ngayTraKetQua)],
        ['Ghi chú', item.ghiChu],
      ];

      const createSequence = () => {
        const sequence = document.createElement('div');
        sequence.className = 'overdue-ticker-sequence';

        overdueList.forEach((item, index) => {
          const entry = document.createElement('span');
          entry.className = 'overdue-ticker-item';
          entry.textContent = fieldsFor(item)
            .map(([label, value]) => `${label}: ${value === null || value === undefined || value === '' ? '-' : value}`)
            .join('  -  ');
          sequence.appendChild(entry);

          if (index < overdueList.length - 1) {
            const separator = document.createElement('span');
            separator.setAttribute('aria-hidden', 'true');
            separator.textContent = '◆';
            sequence.appendChild(separator);
          }
        });

        return sequence;
      };

      const sequence = createSequence();
      const duplicate = sequence.cloneNode(true);
      duplicate.setAttribute('aria-hidden', 'true');
      track.append(sequence, duplicate);
      const duration = Math.max(35, Math.min(180, Math.ceil(sequence.textContent.length / 10)));
      track.style.setProperty('--overdue-ticker-duration', `${duration}s`);
    }

    function checkOverdueNotifications() {
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      
      let overdueList = [];

      currentData.forEach(item => {
      const isPD = (item.daLamVBPhucDap === true || item.daLamVBPhucDap === 1 || item.daLamVBPhucDap === "1" || item.daLamVBPhucDap === "true" || item.daLamVBPhucDap === "Có");
      const isOnline = (item.nopOnline === true || item.nopOnline === 1 || item.nopOnline === "1" || item.nopOnline === "true" || item.nopOnline === "Có");
      
      if (item.ngayTraKetQua && !isPD && !isOnline) {
        const deadline = new Date(item.ngayTraKetQua);
        deadline.setHours(0, 0, 0, 0);
        if (deadline < today) {
          overdueList.push(item);
        }
      }
    });

    updateOverdueTicker(overdueList);

      const badge = document.getElementById('notiBadge');
      const listContainer = document.getElementById('notificationList');
      const timeDisplay = document.getElementById('notiTime');
      
      const now = new Date();
      timeDisplay.innerText = `${now.getHours()}:${String(now.getMinutes()).padStart(2, '0')}`;

      if (overdueList.length > 0) {
        badge.innerText = overdueList.length;
        badge.classList.remove('hidden');

        bellIcon.classList.add('animate-ring'); 
        
        listContainer.innerHTML = '';
        overdueList.forEach(item => {
          const safeSoCV = String(item.soCongVan).replace(/'/g, "\\'");
          const div = document.createElement('div');
          div.className = "p-2 bg-red-50 rounded border border-red-100 text-red-800 flex flex-col gap-1 cursor-pointer hover:bg-red-100 transition shadow-sm";
          div.onclick = function() { jumpToDocument(safeSoCV); };
          
          div.innerHTML = `
            <div class="flex justify-between items-start font-semibold text-xs">
              <span>CV: ${item.soCongVan}</span>
              <span>Hạn: ${item.ngayTraKetQua}</span>
            </div>
            <div class="text-[11px] truncate" title="${item.coQuan}">
              ${item.coQuan}
            </div>
          `;
          listContainer.appendChild(div);
        });
      } else {
        badge.classList.add('hidden');
        listContainer.innerHTML = '<div class="text-center p-4 text-gray-500 text-xs italic">Không có công văn quá hạn</div>';
      }
    }

    function toggleNotificationPanel() {
      const panel = document.getElementById('notificationPanel');
      panel.classList.toggle('hidden');

      document.getElementById('bellIcon').classList.remove('animate-ring');
    }

    function jumpToDocument(soCongVan) {
      document.getElementById('notificationPanel').classList.add('hidden');
      
      const searchInput = document.getElementById('searchSoCongVan');
      searchInput.value = soCongVan;
      
      handleSearchInput();
    
      setTimeout(() => {
        const tableContainer = document.querySelector('.overflow-x-auto');
        if (tableContainer) {
          tableContainer.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }

        const rows = document.querySelectorAll('#tableBody tr');
        rows.forEach(tr => {
          if (tr.innerText.includes(soCongVan)) {
            tr.classList.add('highlight-row');
            setTimeout(() => tr.classList.remove('highlight-row'), 3000);
          }
        });
      }, 300);
    }

    document.addEventListener('click', function(event) {
      const panel = document.getElementById('notificationPanel');
      const button = event.target.closest('button[onclick="toggleNotificationPanel()"]');
      if (!panel.classList.contains('hidden') && !button && !panel.contains(event.target)) {
        panel.classList.add('hidden');
      }
    });

    // ==========================================
    // ĐỒNG BỘ THANH CUỘN NGANG
    // ==========================================
    function syncScrollbars() {
      const mainContainer = document.getElementById('mainTableContainer');
      const topScrollbar = document.getElementById('topScrollbar');
      const topScrollbarContent = document.getElementById('topScrollbarContent');

      if (!mainContainer || !topScrollbar || !topScrollbarContent) return;

      const maxMainScroll = mainContainer.scrollWidth - mainContainer.clientWidth;

      if (maxMainScroll <= 0) {
        topScrollbar.style.display = 'none';
        mainContainer.scrollLeft = 0;
        topScrollbar.scrollLeft = 0;
        return;
      }

      topScrollbar.style.display = 'block';
      topScrollbarContent.style.width = `${topScrollbar.clientWidth + maxMainScroll}px`;
      topScrollbar.scrollLeft = mainContainer.scrollLeft;
    }

    document.addEventListener("DOMContentLoaded", function() {
      const mainContainer = document.getElementById('mainTableContainer');
      const topScrollbar = document.getElementById('topScrollbar');
      const table = document.getElementById('mainTable');

      if (!mainContainer || !topScrollbar || !table) return;

      topScrollbar.addEventListener('scroll', function() {
        mainContainer.scrollLeft = topScrollbar.scrollLeft;
      });

      mainContainer.addEventListener('scroll', function() {
        topScrollbar.scrollLeft = mainContainer.scrollLeft;
      });

      window.addEventListener('resize', syncScrollbars);
      if (window.ResizeObserver) {
        const scrollbarResizeObserver = new ResizeObserver(syncScrollbars);
        scrollbarResizeObserver.observe(mainContainer);
        scrollbarResizeObserver.observe(table);
      }
      if (window.MutationObserver) {
        const tableMutationObserver = new MutationObserver(syncScrollbars);
        tableMutationObserver.observe(table, {
          childList: true,
          subtree: true,
          characterData: true
        });
      }

      setTimeout(syncScrollbars, 500);
    });

    function autoFillMaxRow(index, soCongVan) {
      const row = currentData[index];
      if (!row) return;

      const total = (Number(row.soToChuc) || 0) + (Number(row.soCaNhan) || 0);
      
      if (total === 0) {
        showToast("Dòng này không có số lượng tổ chức hoặc cá nhân!", "error");
        return;
      }

      const htInput = document.getElementById(`row_ht_${index}`);
      if (htInput && !htInput.disabled) {
        htInput.value = total;
      }

      const tnInput = document.getElementById(`row_tn_${index}`);
      const sxInput = document.getElementById(`row_sx_${index}`);
      
      if (tnInput && !tnInput.disabled) tnInput.value = 0;
      if (sxInput && !sxInput.disabled) sxInput.value = 0;

      showToast(`Đã chọn tất cả thành công!`, "success");
    }

    function handleCheckboxClick(checkbox, index, soCongVan) {
      const row = currentData[index];
      if (!row) return;

      const htInput = document.getElementById(`row_ht_${index}`);
      const htOnScreen = Number(htInput.value) || 0; 
      
      const soTC = Number(row.soToChuc) || 0;
      const soCN = Number(row.soCaNhan) || 0;
      const totalRequired = soTC + soCN;

      if (checkbox.checked && htOnScreen !== totalRequired) {
        checkbox.checked = false;
        
        alert(`KHÔNG THỂ HOÀN THÀNH!\n\nVui lòng bấm nút "Cập nhật" sau khi gõ đủ số lượng!`);
        
        htInput.classList.add('border-red-500', 'ring-2', 'ring-red-300');
        htInput.focus();
        setTimeout(() => htInput.classList.remove('border-red-500', 'ring-2', 'ring-red-300'), 2000);
        return;
      }

      const tr = checkbox.closest('tr');
      
      const allInputs = tr.querySelectorAll('input');
      
      if (checkbox.checked) {
        allInputs.forEach(inp => {
          if (inp !== checkbox) {
            inp.disabled = true;
            inp.classList.add('bg-gray-100', 'text-gray-500', 'cursor-not-allowed');
            if (inp.type === 'checkbox') {
              inp.classList.add('opacity-50');
            }
          }
        });
      } else {
        allInputs.forEach(inp => {
          inp.disabled = false;
          inp.classList.remove('bg-gray-100', 'text-gray-500', 'cursor-not-allowed');
          if (inp.type === 'checkbox') {
            inp.classList.remove('opacity-50');
          }
        });
      }
    }
    
    function handleNopOnlineClick(checkbox, index, soCongVan) {
        const row = currentData[index];
        if (!row) return;

        const htInput = document.getElementById(`row_ht_${index}`);
        const htOnScreen = Number(htInput.value) || 0; 
        
        const soTC = Number(row.soToChuc) || 0;
        const soCN = Number(row.soCaNhan) || 0;
        const totalRequired = soTC + soCN;

        if (checkbox.checked && htOnScreen !== totalRequired) {
            checkbox.checked = false;
            
            alert(`KHÔNG THỂ HOÀN THÀNH!\n\n\nVui lòng bấm nút "Cập nhật" sau khi gõ đủ số lượng!`);
            
            htInput.classList.add('border-red-500', 'ring-2', 'ring-red-300');
            htInput.focus();
            setTimeout(() => htInput.classList.remove('border-red-500', 'ring-2', 'ring-red-300'), 2000);
            return;
        }

        const tr = checkbox.closest('tr');
        const allInputs = tr.querySelectorAll('input');
        const isPhucDapChecked = document.getElementById(`row_pd_${index}`).checked;

        if (checkbox.checked) {
            allInputs.forEach(inp => {
                if (inp !== checkbox) {
                    inp.disabled = true;
                    inp.classList.add('bg-gray-100', 'text-gray-500', 'cursor-not-allowed');
                    if (inp.type === 'checkbox') inp.classList.add('opacity-50');
                }
            });
        } else {
            if (!isPhucDapChecked) {
                allInputs.forEach(inp => {
                    inp.disabled = false;
                    inp.classList.remove('bg-gray-100', 'text-gray-500', 'cursor-not-allowed');
                    if (inp.type === 'checkbox') inp.classList.remove('opacity-50');
                });
            }
        }
    }

    function handleCoQuanChange() {
      const coQuanInput = document.getElementById('inputCoQuan');
      const mailInput = document.getElementById('inputMailDauMoi');
      const donViSelect = document.getElementById('inputDonVi'); 
      
      const selectedCoQuan = coQuanInput.value.trim().toLowerCase().replace(/\s+/g, ' ');
      
      donViSelect.innerHTML = '<option value="">-- Chọn đơn vị --</option>'; 
      mailInput.value = '';

      if (!selectedCoQuan || !danhSachCoQuanData || danhSachCoQuanData.length === 0) {
        return;
      }
      
      const foundItems = danhSachCoQuanData.filter(item => {
        const tenCQ = (item.tenCoQuan || item.TenCoQuan || "").trim().toLowerCase().replace(/\s+/g, ' ');
        return tenCQ === selectedCoQuan;
      });
      
      if (foundItems.length > 0) {
        let hasCap2 = false; 

        foundItems.forEach(item => {
          const cap2 = item.donViCap2 || ""; 
          if (cap2.trim() !== "") { 
            hasCap2 = true;
            const option = document.createElement('option');
            option.value = cap2;
            option.textContent = cap2;
            donViSelect.appendChild(option);
          }
        });

        const itemKhongCap2 = foundItems.find(item => (item.donViCap2 || "").trim() === "");
        
        if (itemKhongCap2 && itemKhongCap2.mailCongVu) {
          mailInput.value = itemKhongCap2.mailCongVu;
        } else if (!hasCap2) {
          mailInput.value = foundItems[0].mailCongVu || "";
        }
        
        if (mailInput.value) {
          mailInput.classList.add('bg-emerald-50', 'border-emerald-300');
          setTimeout(() => mailInput.classList.remove('bg-emerald-50', 'border-emerald-300'), 1500);
        }
      }

      if (mailInput.value) {
        document.getElementById('btnSaveMail').classList.remove('hidden');
      }
    }

    function handleDonViChange() {
      const coQuanInput = document.getElementById('inputCoQuan');
      const donViSelect = document.getElementById('inputDonVi');
      const mailInput = document.getElementById('inputMailDauMoi');
      const btnSaveMail = document.getElementById('btnSaveMail');
      
      const selectedCoQuan = coQuanInput.value.trim().toLowerCase().replace(/\s+/g, ' ');
      const selectedDonVi = donViSelect.value.trim().toLowerCase();
      
      if (!selectedDonVi) {
        mailInput.value = '';
        if(btnSaveMail) btnSaveMail.classList.add('hidden');
        return;
      }

      const found = danhSachCoQuanData.find(item => {
        const tenCQ = (item.tenCoQuan || item.TenCoQuan || "").trim().toLowerCase().replace(/\s+/g, ' ');
        const cap2 = (item.donViCap2 || "").trim().toLowerCase();
        return tenCQ === selectedCoQuan && cap2 === selectedDonVi;
      });
      
      if (found) {
        const mail = found.mailCongVu || ""; 
        if (mail.trim() !== "") {
          mailInput.value = mail;
          mailInput.classList.add('bg-emerald-50', 'border-emerald-300');
          setTimeout(() => mailInput.classList.remove('bg-emerald-50', 'border-emerald-300'), 1500);
        } else {
          mailInput.value = '';
        }
      } else {
        mailInput.value = '';
      }

      if(btnSaveMail) btnSaveMail.classList.remove('hidden');
    }

    function toggleClearCoQuanBtn() {
      const input = document.getElementById('inputCoQuan');
      const btn = document.getElementById('btnClearCoQuan');
      if (input.value.length > 0) {
        btn.classList.remove('hidden');
      } else {
        btn.classList.add('hidden');
      }
    }

    function clearCoQuanInput() {
      const input = document.getElementById('inputCoQuan');
      input.value = '';
      document.getElementById('btnClearCoQuan').classList.add('hidden');
      
      handleCoQuanChange(); 
      
      input.focus();
    }
    
    function toggleClearDonViBtn() {
      const select = document.getElementById('inputDonVi');
      const btn = document.getElementById('btnClearDonVi');
      if (select.value !== "") {
        btn.classList.remove('hidden');
      } else {
        btn.classList.add('hidden');
      }
    }

    function clearDonViInput() {
      const select = document.getElementById('inputDonVi');
      select.value = "";
      document.getElementById('btnClearDonVi').classList.add('hidden');
      
      document.getElementById('inputMailDauMoi').value = '';
      document.getElementById('btnSaveMail').classList.add('hidden');
    }

    // ==========================================
    // XỬ LÝ LỌC THÁNG, QUÁ HẠN & XUẤT EXCEL
    // ==========================================
    let isOverdueFilterActive = false;
    let isUnrepliedFilterActive = false;

    function toggleOverdueFilter() {
      isOverdueFilterActive = !isOverdueFilterActive;
      const btn = document.getElementById('btnFilterOverdue');
      const alertBox = document.getElementById('overdueAlert');
      
      if (isOverdueFilterActive) {
        btn.className = "px-3 py-1.5 bg-orange-600 text-white border border-orange-700 rounded-lg text-sm font-medium hover:bg-orange-700 transition shadow-inner";
        alertBox.classList.remove('hidden');
      } else {
        btn.className = "px-3 py-1.5 bg-orange-100 text-orange-700 border border-orange-200 rounded-lg text-sm font-medium hover:bg-orange-200 transition";
        alertBox.classList.add('hidden');
      }
      currentPage = 1;
      filterAndRender();
    }

    function toggleUnrepliedFilter() {
      isUnrepliedFilterActive = !isUnrepliedFilterActive;
      const btn = document.getElementById('btnFilterUnreplied');
      const alertBox = document.getElementById('unrepliedAlert');
      
      if (isUnrepliedFilterActive) {
        btn.className = "px-3 py-1.5 bg-blue-600 text-white border border-blue-700 rounded-lg text-sm font-medium hover:bg-blue-700 transition shadow-inner";
        alertBox.classList.remove('hidden');
      } else {
        btn.className = "px-3 py-1.5 bg-blue-100 text-blue-700 border border-blue-200 rounded-lg text-sm font-medium hover:bg-blue-200 transition";
        alertBox.classList.add('hidden');
      }
      currentPage = 1;
      filterAndRender();
    }

    function exportToExcel() {
      const exportData = AppBusiness.filterRecords(currentData, getCurrentFilters());
      if (exportData.length === 0) {
        showToast("Không có dữ liệu để xuất!", "error");
        return;
      }

      let csvContent = "\uFEFF";
      csvContent += "STT;S? C?ng V?n;C? Quan\n";
      exportData.forEach((row, index) => {
        const soCV = String(row.soCongVan || '').replace(/;/g, '');
        const coQuan = String(row.coQuan || '').replace(/;/g, '');
        csvContent += (index + 1) + ";" + soCV + ";" + coQuan + "\n";
      });

      const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
      const link = document.createElement("a");
      const url = URL.createObjectURL(blob);
      link.setAttribute("href", url);
      link.setAttribute("download", "Danh_Sach_Cong_Van_" + new Date().toLocaleDateString('vi-VN').replace(/\//g, '-') + ".csv");
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      showToast("Đã xuất file Excel thành công!", "success");
    }

    async function submitNewDonVi() {
      const coQuan = document.getElementById('addCoQuanInput').value.trim();
      const donVi = document.getElementById('addDonViInput').value.trim();
      const mail = document.getElementById('addMailInput').value.trim();
      const btn = document.getElementById('btnSubmitDonVi');

      if (!donVi) {
        showToast("Vui lòng nhập tên Đơn vị (Cấp 2)!", "error");
        return;
      }

      const isDuplicate = AppBusiness.isDuplicateOrganization(danhSachCoQuanData, coQuan, donVi);

      if (isDuplicate) {
        showToast(`Đơn vị "${donVi}" đã tồn tại trong Cơ quan này!`, "error");
        
        const inputDonVi = document.getElementById('addDonViInput');
        inputDonVi.classList.add('border-red-500', 'ring-2', 'ring-red-300');
        inputDonVi.focus();
        setTimeout(() => inputDonVi.classList.remove('border-red-500', 'ring-2', 'ring-red-300'), 2000);
        return;
      }

      btn.disabled = true;
      btn.innerText = "Đang lưu...";

      const payload = {
        action: "addDanhSachCoQuan", 
        tenCoQuan: coQuan,
        donViCap2: donVi,
        mailCongVu: mail
      };

      try {
        await AppData.addOrganization(payload);

        showToast("Đã thêm Đơn vị mới thành công!", "success");
        closeAddDonViModal();
        
        await loadDanhSachCoQuan(); 
        
        setTimeout(() => {
          const select = document.getElementById('inputDonVi');
          for (let i = 0; i < select.options.length; i++) {
            if (select.options[i].value === donVi) {
              select.selectedIndex = i;
              break;
            }
          }
          if (mail) {
            document.getElementById('inputMailDauMoi').value = mail;
          }
        }, 1000);

      } catch (err) {
        console.error("Lỗi khi thêm đơn vị:", err);
        showToast("Lỗi kết nối, vui lòng thử lại!", "error");
      } finally {
        btn.disabled = false;
        btn.innerText = "Lưu vào hệ thống";
      }
    }

    let allDonViSuggestions = [];

    function openAddDonViModal() {
      const coQuanHienTai = document.getElementById('inputCoQuan').value.trim();
      if (!coQuanHienTai) {
        showToast("Vui lòng chọn Cơ quan (Cấp 1) trước khi thêm Đơn vị!", "error");
        return;
      }
      
      document.getElementById('addCoQuanInput').value = coQuanHienTai;
      document.getElementById('addDonViInput').value = '';
      document.getElementById('addMailInput').value = '';
      document.getElementById('suggestList').classList.add('hidden');
      
      allDonViSuggestions = [];
      if (danhSachCoQuanData && danhSachCoQuanData.length > 0) {
        const currentCoQuanNormalized = coQuanHienTai.toLowerCase().replace(/\s+/g, ' ');
        
        danhSachCoQuanData.forEach(item => {
          const tenCQ = (item.tenCoQuan || item.TenCoQuan || "").trim().toLowerCase().replace(/\s+/g, ' ');
          const cap2 = (item.donViCap2 || "").trim();
          
          if (tenCQ === currentCoQuanNormalized && cap2 !== "") {
            allDonViSuggestions.push(cap2);
          }
        });
      }
      
      renderSuggestList(allDonViSuggestions);

      const modal = document.getElementById('addDonViModal');
      modal.classList.remove('hidden');
      modal.classList.add('flex');
    }

    function toggleSuggestList(show) {
      const list = document.getElementById('suggestList');
      if (show) {
        list.classList.remove('hidden');
        filterSuggestList();
      } else {
        list.classList.add('hidden');
      }
    }

    function filterSuggestList() {
      const input = document.getElementById('addDonViInput').value.toLowerCase();
      const filtered = allDonViSuggestions.filter(dv => dv.toLowerCase().includes(input));
      renderSuggestList(filtered);
      document.getElementById('suggestList').classList.remove('hidden');
    }

    function renderSuggestList(items) {
      const list = document.getElementById('suggestList');
      list.innerHTML = '';
      
      if (items.length === 0) {
        list.innerHTML = '<div class="px-3 py-2 text-sm text-gray-400 italic">Chưa có đơn vị nào...</div>';
        return;
      }

      items.forEach(item => {
        const div = document.createElement('div');
        div.className = "px-3 py-2 text-sm text-gray-700 hover:bg-indigo-50 cursor-pointer transition border-b border-gray-50 last:border-0";
        div.innerText = item;
        div.onclick = function() {
          document.getElementById('addDonViInput').value = item;
          toggleSuggestList(false);
        };
        list.appendChild(div);
      });
    }

    document.addEventListener('click', function(event) {
      const input = document.getElementById('addDonViInput');
      const list = document.getElementById('suggestList');
      if (input && list && !input.contains(event.target) && !list.contains(event.target)) {
        list.classList.add('hidden');
      }
    });
    
    function closeAddDonViModal() {
      const modal = document.getElementById('addDonViModal');
      modal.classList.remove('flex');
      modal.classList.add('hidden');
    }
    
    async function updateMailToSheet() {
      const coQuan = document.getElementById('inputCoQuan').value.trim();
      const donVi = document.getElementById('inputDonVi').value.trim();
      const mail = document.getElementById('inputMailDauMoi').value.trim();
      const btn = document.getElementById('btnSaveMail');

      if (!coQuan) {
        showToast("Vui lòng chọn Cơ quan!", "error");
        return;
      }

      btn.innerText = "Đang lưu...";
      btn.disabled = true;

      const payload = {
        action: "updateMailDanhSach",
        tenCoQuan: coQuan,
        donViCap2: donVi,
        mailCongVu: mail
      };

      try {
        await AppData.updateOrganizationMail(payload);

        showToast("Đã cập nhật Mail thành công!", "success");
        btn.classList.add('hidden');
        
        await loadDanhSachCoQuan(); 

      } catch (err) {
        console.error("Lỗi cập nhật mail:", err);
        showToast("Cập nhật thất bại!", "error");
      } finally {
        btn.innerText = "Lưu Mail";
        btn.disabled = false;
      }
    }

    function resizeTrichYeuInput() {
      const input = document.getElementById('inputTrichYeu');
      input.style.height = 'auto';
      const borderHeight = input.offsetHeight - input.clientHeight;
      input.style.height = `${input.scrollHeight + borderHeight}px`;
    }

    function showTrichYeuSuggest() {
      const input = document.getElementById('inputTrichYeu');
      const listContainer = document.getElementById('trichYeuSuggestList');
      const keyword = input.value.trim().toLowerCase();

      const seenTrichYeu = new Set();
      let uniqueTrichYeu = currentData.reduce((suggestions, item) => {
        const text = String(item.trichYeu || '')
          .replace(/[\u200B-\u200D\uFEFF]/g, '')
          .replace(/\s+/g, ' ')
          .trim()
          .normalize('NFC');
        const key = text.toLowerCase();
        if (text && !seenTrichYeu.has(key)) {
          seenTrichYeu.add(key);
          suggestions.push(text);
        }
        return suggestions;
      }, []);
      
      if (keyword) {
        uniqueTrichYeu = uniqueTrichYeu.filter(t => t.toLowerCase().includes(keyword));
      }
      
      if (uniqueTrichYeu.length === 0) {
        listContainer.classList.add('hidden');
        return;
      }

      listContainer.innerHTML = '';
      uniqueTrichYeu.slice(0, 10).forEach(text => {
        const div = document.createElement('div');
        div.className = "flex justify-between items-center px-3 py-2 text-sm text-gray-700 hover:bg-indigo-50 border-b border-gray-50 cursor-pointer";
        
        const span = document.createElement('span');
        span.innerText = text;
        span.className = "truncate pr-2 flex-1";
        span.onclick = function() {
          input.value = text;
          resizeTrichYeuInput();
          listContainer.classList.add('hidden');
        };

        const btnX = document.createElement('button');
        btnX.innerHTML = '✕';
        btnX.className = "text-red-400 hover:text-red-600 font-bold px-2";
        btnX.type = "button";
        btnX.onclick = function(e) {
          e.stopPropagation();
          div.remove();
        };

        div.appendChild(span);
        div.appendChild(btnX);
        listContainer.appendChild(div);
      });
      
      listContainer.classList.remove('hidden');
    }

    function toggleClearTrichYeuBtn() {
      const input = document.getElementById('inputTrichYeu');
      const button = document.getElementById('btnClearTrichYeu');
      button.classList.toggle('hidden', input.value.length === 0);
    }

    function clearTrichYeuInput() {
      const input = document.getElementById('inputTrichYeu');
      const listContainer = document.getElementById('trichYeuSuggestList');
      input.value = '';
      resizeTrichYeuInput();
      toggleClearTrichYeuBtn();
      listContainer.classList.add('hidden');
      input.focus();
      listContainer.classList.add('hidden');
    }

    document.addEventListener('click', function(event) {
      const input = document.getElementById('inputTrichYeu');
      const list = document.getElementById('trichYeuSuggestList');
      if (input && list && !input.contains(event.target) && !list.contains(event.target)) {
        list.classList.add('hidden');
      }
    });
