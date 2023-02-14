var manahej = {
  1: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10],
  2: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20],
  3: [
    1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21,
    22, 23, 24, 25, 26, 27, 28, 29, 30,
  ],
};

const what_could_be_converted_to = (prev_manhaj, prev_level) => {
  if (prev_manhaj == 4) {
    return console.log(
      `${prev_manhaj} - ${prev_level} can be converted to: 4 - ${
        prev_level + 1
      } or 6 - ${prev_level * 2 + 1}`
    );
  } else if (prev_manhaj == 6) {
    if (prev_level % 2 != 0) {
      return console.log(
        `${prev_manhaj} - ${prev_level} can't be converted.. try on next level`
      );
    }
    return console.log(
      `${prev_manhaj} - ${prev_level} can be converted to: 4 - ${
        prev_level / 2 + 1
      } or 6 - ${prev_level + 1}`
    );
  }

  if (prev_level % prev_manhaj != 0) {
    return console.log(
      `${prev_manhaj} - ${prev_level} can't be converted.. try on next level`
    );
  }

  return console.log(
    `${prev_manhaj} - ${prev_level} can be converted to: '1 - ${
      (prev_level / prev_manhaj) * 1 + 1
    } or 2 - ${(prev_level / prev_manhaj) * 2 + 1} or 3 - ${
      (prev_level / prev_manhaj) * 3 + 1
    }`
  );
};

what_could_be_converted_to(2, 4);
