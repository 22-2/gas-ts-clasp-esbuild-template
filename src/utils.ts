
/**
 * argがnull/undefinedのときは例外を送出し、そうでなければcallable(arg)の結果を返します
 */
export function orThrow<T, U = void>(
  arg: T | null | undefined,
  callable: (arg: T) => U,
  opts?: { message?: string; }
): U {
  // eslint-disable-line
  if (arg != null) {
    // eslint-disable-line
    return callable(arg);
  }
  throw Error(
    opts?.message ?? '予期せずundefined/nullが発生しましたので中断しました'
  );
}

export function parseMd(markdownText: string): {
  frontmatter: string;
  content: string;
} {
  const frontmatterPattern = /^---\n(.*?)\n---\n(.*)/s;
  const match = markdownText.match(frontmatterPattern);

  return orThrow(match, match => ({
    frontmatter: match[1],
    content: match[2],
  }));
}

// eslint-disable-line
export function dev(...args: any[]) {
  Logger.log(args);
}
