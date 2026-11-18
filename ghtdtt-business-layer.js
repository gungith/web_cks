(function () {
  function isTrueFlag(value) {
    return value === true || value === 1 || value === "1" || value === "true" || value === "Có";
  }

  function isReplied(item) {
    return isTrueFlag(item.daLamVBPhucDap) || isTrueFlag(item.nopOnline);
  }

  function hasDocumentNumber(item) {
    return String(item.soCongVan || "").trim() !== "";
  }

  function isOverdue(item, today = new Date()) {
    if (!item.ngayTraKetQua || isReplied(item)) return false;

    const todayStart = new Date(today);
    todayStart.setHours(0, 0, 0, 0);
    const deadline = new Date(item.ngayTraKetQua);
    deadline.setHours(0, 0, 0, 0);
    return deadline < todayStart;
  }

  function filterRecords(records, filters) {
    const today = new Date(filters.today || new Date());
    today.setHours(0, 0, 0, 0);
    const hasSearchQuery = Boolean(filters.searchQuery);

    return records.filter(item => {
      if (!item.ngay && !hasSearchQuery) return false;

      let matchDate = true;

      if (!hasSearchQuery) {
        const itemDate = new Date(item.ngay);
        itemDate.setHours(0, 0, 0, 0);

        if (filters.filterType === "today") {
          matchDate = itemDate.getTime() === today.getTime();
        } else if (filters.filterType === "week") {
          const currentDayOfWeek = today.getDay() === 0 ? 6 : today.getDay() - 1;
          const startOfWeek = new Date(today);
          startOfWeek.setDate(today.getDate() - currentDayOfWeek);
          const endOfWeek = new Date(startOfWeek);
          endOfWeek.setDate(startOfWeek.getDate() + 6);
          matchDate = itemDate >= startOfWeek && itemDate <= endOfWeek;
        } else if (filters.filterType === "month") {
          matchDate = itemDate.getFullYear() === today.getFullYear()
            && itemDate.getMonth() === today.getMonth();
        } else if (filters.filterType === "custom" && filters.customDate) {
          matchDate = item.ngay === filters.customDate;
        } else if (filters.filterType === "range") {
          if (filters.startDate && filters.endDate) {
            matchDate = item.ngay >= filters.startDate && item.ngay <= filters.endDate;
          } else if (filters.startDate) {
            matchDate = item.ngay >= filters.startDate;
          } else if (filters.endDate) {
            matchDate = item.ngay <= filters.endDate;
          }
        }
      }

      const matchSearch = !filters.searchQuery
        || String(item.soCongVan || "").toLowerCase().includes(filters.searchQuery);
      const matchOverdue = !filters.overdueOnly || isOverdue(item, today);
      const matchUnreplied = !filters.unrepliedOnly || !isReplied(item);

      return matchDate && matchSearch && matchOverdue && matchUnreplied;
    });
  }

  function calculateKpis(records) {
    return records.reduce((totals, item) => {
      if (hasDocumentNumber(item)) totals.soVanBan += 1;
      totals.soToChuc += Number(item.soToChuc) || 0;
      totals.soCaNhan += Number(item.soCaNhan) || 0;
      totals.tongTiepNhan += Number(item.tongTiepNhan) || 0;
      totals.tongSanXuat += Number(item.tongSanXuat) || 0;
      totals.hoanThanh += Number(item.hoanThanh) || 0;
      if (isReplied(item)) totals.soCongVanHoanThanh += 1;
      return totals;
    }, {
      soVanBan: 0,
      soToChuc: 0,
      soCaNhan: 0,
      tongTiepNhan: 0,
      tongSanXuat: 0,
      hoanThanh: 0,
      soCongVanHoanThanh: 0
    });
  }

  function calculateDocumentCompletion(records) {
    const documents = records.filter(hasDocumentNumber);
    const total = documents.length;
    const completed = documents.filter(isReplied).length;

    return {
      total,
      completed,
      percentage: total ? Math.round((completed / total) * 100) : 0
    };
  }

  function isDuplicateOrganization(organizations, coQuan, donVi) {
    const normalize = value => String(value || "").trim().toLowerCase().replace(/\s+/g, " ");
    const normalizedCoQuan = normalize(coQuan);
    const normalizedDonVi = normalize(donVi);

    return organizations.some(item =>
      normalize(item.tenCoQuan || item.TenCoQuan) === normalizedCoQuan
      && normalize(item.donViCap2) === normalizedDonVi
    );
  }

  function createNewRecord(values) {
    const coQuan = values.donVi
      ? `${values.donVi} ${values.coQuan}`.trim().replace(/\s+/g, " ")
      : values.coQuan;
    const total = values.soToChuc + values.soCaNhan;

    return {
      ngay: values.ngay,
      ngayCongVan: values.ngayCongVan,
      ngayTraKetQua: values.ngayTraKetQua,
      soCongVan: values.soCongVan,
      coQuan,
      trichYeu: values.trichYeu,
      soToChuc: values.soToChuc,
      soCaNhan: values.soCaNhan,
      mailDauMoi: values.mailDauMoi,
      tongTiepNhan: total,
      tongSanXuat: 0,
      hoanThanh: 0,
      daXuLy: true
    };
  }

  window.AppBusiness = {
    isTrueFlag,
    isReplied,
    isOverdue,
    filterRecords,
    calculateKpis,
    calculateDocumentCompletion,
    isDuplicateOrganization,
    createNewRecord
  };
})();
