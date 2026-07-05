# email-to-print

My family emails a PDF or a photo to a print address on our domain. A small
Python script watches that mailbox, checks the sender against a list of four
people, converts the file if it needs converting, and hands it to CUPS. The
printer is a boring HP laser from 2017. That's the whole system.

A bunch of people asked for a writeup after I posted about this, so this
README is the writeup. Fair warning that none of it is clever. It's mostly
plumbing, and half of it is a mail filter.

The reason it exists is dumber than the system. My wife would airdrop me a
permission slip, I'd open it on my desktop, print it, and wonder why I was
involved at all. The printer vendors solve this with their cloud print
services, but I didn't want school forms bouncing through someone else's
servers, and those services have a habit of getting discontinued.

This is the actual code running in my house, with my personal values
stripped out. It's been printing homework and permission slips since spring
2026.

## The pieces

```
phone/laptop
     |  email with attachment
     v
your mail provider  ->  (ProtonMail Bridge, if Proton)  ->  IMAP folder
                                                              |
                                                    print-poller (this repo)
                                                              |
                                        allowlist check -> convert -> lp
                                                              |
                                                        CUPS -> printer
```

The mailbox is a Proton account I already had. Proton doesn't do plain IMAP,
so ProtonMail Bridge runs in Docker and exposes the account as IMAP on
localhost. The bridge uses a self-signed cert, which is why the example env
sets `TLS_VERIFY=false`; that setting is for localhost only, and verification
stays on for anything else. If you use Gmail or anything normal you can skip
the bridge entirely and point straight at your provider. The compose pins
the bridge image to the exact digest running in my house; it's a community
image, so read the comment above it before trusting it with credentials.

A filter at the mail provider moves anything addressed to the print alias
into its own folder, so the script never touches the real inbox. That filter
is doing more work than any of my code.

The script polls that folder once a minute:

- If the sender isn't on `ALLOWED_SENDERS`, the message moves to a rejected
  folder and the sender gets a reply saying so. An empty allowlist means
  nothing prints. I made it fail closed after about ten seconds of imagining
  what happens if it didn't.
- PDFs and images go straight to `lp`. Word docs and spreadsheets get
  converted to PDF first with headless LibreOffice, which works maybe 95% of
  the time and produces something ugly but printable the rest.
- If there's no printable attachment, it renders the email body itself and
  prints that, so forwarding an email prints the email. That one was an
  afterthought and it's become the most-used feature in the house.
- Every handled message gets moved out of the watch folder, printed or
  rejected. That move is the entire dedup system. There's no database. If
  the folder has a message in it, it hasn't been dealt with yet. Strictly
  speaking that's at-least-once: if the process died between printing and
  moving, you'd get a duplicate. I expected to regret this and haven't.
- Allowlisted senders get a confirmation email whether the print worked or
  not, unless you enable auth checks and the message fails them. It sounds
  like a gimmick until someone prints from the grocery
  store and wants to know if it worked. Rejected strangers get silence by
  default, on purpose. Their From address is unverified, and replying to
  spoofed mail is how you become a backscatter cannon.

Stdlib Python only. The container's external dependencies are `lp`
(cups-client) and `soffice` (libreoffice-nogui). There's a JSON health
endpoint on `127.0.0.1:2631/health` if you monitor things; it stays on
localhost unless you rebind it.

## Setup

1. Get your printer working in CUPS on the Docker host first. `lpstat -p`
   should list a queue and `lp -d Your_Queue test.pdf` should print. If that
   doesn't work, nothing below will either.
2. `cp .env.example .env` and fill it in. `ALLOWED_SENDERS` and `PRINT_TO`
   are the two that matter.
3. Create the filter at your mail provider that routes `PRINT_TO` mail into
   a dedicated folder, and set `SOURCE_FOLDER` to it.
4. If you're on Proton: `docker compose run --rm protonmail-bridge init`
   once to log in interactively. Anyone else: delete that service from the
   compose file and point `IMAP_HOST`/`IMAP_PORT` at your provider.
5. `docker compose up -d --build`
6. Send yourself a test with `DRY_RUN=true` first. The logs show what would
   have printed. Then flip it off.

## The one honest weakness

The allowlist checks the From header, and From headers can be faked.
Somebody who knew the print address and an allowlisted address could print
at my house. I thought about requiring SPF/DKIM checks and decided the worst
case is wasted paper, so I left it off. It's there if your threat model
disagrees: set `REQUIRE_AUTH_PASS=true` and the script also requires an SPF
or DKIM pass in the Authentication-Results header. Auth failures reject
silently, no confirmation, same backscatter logic as strangers. Fair warning that this
just reads the header your provider stamped, and headers can be forged too.
It raises the bar, it is not a bouncer. Keeping the print address off the
public internet helps more than either. Mine only exists in family address
books.

`MAX_ATTACH_MB` caps individual attachment size so nobody mails you a 200MB
scan by accident. It's a byte cap, not a page cap. A slim 400 page PDF will
still print 400 pages, so don't allowlist anyone you wouldn't trust with
your toner.

## Things that have actually gone wrong

A few months of running it, in order of occurrence:

- The Proton bridge needed a re-login twice after restarts. The poller just
  logs IMAP failures and retries every cycle, so fixing the bridge is the
  whole fix. It catches up on its own.
- LibreOffice made a mess of a fancy resume template once. The confirmation
  reply meant we at least knew it printed before walking to the printer.
- The printer being off turned out not to matter. CUPS queues the job and
  prints it when the printer comes back on, which confused my son for a
  solid day.

## Questions people actually asked

**Why not HP ePrint / Epson Connect / the vendor cloud thing?**
It routes your documents through the vendor's servers, needs an account, and
has been discontinued or broken for various models over the years. This runs
on my hardware and outlives any vendor's product decisions.

**My printer has a built-in email client (some Kyocera, Ricoh, etc). Why
not that?**
If yours does and you trust its firmware with your mail credentials, use it,
genuinely. Most home printers don't have one, the ones that do rarely
support modern auth, and a firmware update can take the feature away. This
also gets you the allowlist, size caps, Office conversion, and
confirmations, which the built-in clients mostly don't.

**Can I use Gmail instead of Proton?**
Yes, no code changes. Skip the bridge and set `IMAP_HOST=imap.gmail.com`,
`IMAP_PORT=993`, `IMAP_SSL=true`, `TLS_VERIFY=true`, and an app password.
Use `you+print@gmail.com` with a filter as the print address.

**What about the scan-to-email then email-to-print infinite loop?**
Don't put your scanner's outgoing address on the allowlist. That's the whole
fix. Someone spotted this in the original thread immediately and honestly it
would have gotten me.

**I'm not in the US and letter paper is wrong.**
Set `MEDIA=a4` in your `.env`. `SIDES=two-sided-long-edge` gets you duplex
while you're in there.

## License

MIT. If it eats your homework, that's on you, but open an issue anyway.
