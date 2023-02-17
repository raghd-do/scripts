const convert_to = (prev_manhaj, prev_level) => {
  var manhaj_list = [];

  if (prev_manhaj === 4) {
    manhaj_list.push({ m: 4, l: prev_level + 1 });
    manhaj_list.push({ m: 6, l: prev_level * 2 + 1 });
    return manhaj_list;
  } else if (prev_manhaj === 6) {
    if (prev_level % 2 !== 0) {
      manhaj_list.push({ m: 6, l: prev_level + 1 });
      return manhaj_list;
    }
    manhaj_list.push({ m: 4, l: prev_level / 2 + 1 });
    manhaj_list.push({ m: 6, l: prev_level + 1 });
    return manhaj_list;
  }

  if (prev_level % prev_manhaj !== 0) {
    manhaj_list.push({ m: prev_manhaj, l: prev_level + 1 });
    return manhaj_list;
  }

  manhaj_list.push({ m: 1, l: (prev_level / prev_manhaj) * 1 + 1 });
  manhaj_list.push({ m: 2, l: (prev_level / prev_manhaj) * 2 + 1 });
  manhaj_list.push({ m: 3, l: (prev_level / prev_manhaj) * 3 + 1 });
  return manhaj_list;
};
