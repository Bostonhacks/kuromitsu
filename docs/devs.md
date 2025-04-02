This requires the use of a Google OAuth Client.

For BostonHacks, this Client is created under `bostonhackstechteam@gmail.com` under the `kuromitsu` project.

## Concurrency
Behind the scenes, this uses python's concurrent.futures.ThreadPoolExecutor to create worker threads to take on sending each email. This isn't a true parallel execution but it keeps the CPU busy while the program waits for a response from Google. Threading is used here as opposed to subprocesses since this is a primarily I/O bound task (network calls to Google) and there are shared data structures. 